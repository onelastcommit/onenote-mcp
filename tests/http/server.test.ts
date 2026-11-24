import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  HTTP_HOST_ENV,
  HTTP_PORT_ENV,
  HTTP_TOKEN_ENV,
  getHttpHost,
  getHttpPort,
  getHttpToken,
} from '@/config.js';
import { createHttpServer, startHttpServer, type RunningHttpServer } from '@/http/server.js';
import { createServer as createMcpServer } from '@/index.js';
import type * as McpIndex from '@/index.js';

vi.mock('../../src/auth/index.js', () => ({
  getAccessToken: vi.fn(async () => 'fake-access-token'),
}));

vi.mock('../../src/index.js', async (importOriginal) => {
  const actual = await importOriginal<typeof McpIndex>();
  return { ...actual, createServer: vi.fn(actual.createServer) };
});

const TOKEN = 'test-bearer-token-do-not-use';

const initializePayload = {
  jsonrpc: '2.0',
  id: 1,
  method: 'initialize',
  params: {
    protocolVersion: '2025-03-26',
    capabilities: {},
    clientInfo: { name: 'vitest', version: '0.0.0' },
  },
};

const parseMcpResponseBody = async (res: Response): Promise<unknown> => {
  const text = await res.text();
  const contentType = res.headers.get('content-type') ?? '';
  if (contentType.includes('text/event-stream')) {
    for (const line of text.split('\n')) {
      if (line.startsWith('data: ')) return JSON.parse(line.slice('data: '.length));
    }
    throw new Error(`SSE body had no data frame: ${text}`);
  }
  return JSON.parse(text);
};

const mcpHeaders = (token: string | undefined): Record<string, string> => {
  const h: Record<string, string> = {
    'Content-Type': 'application/json',
    Accept: 'application/json, text/event-stream',
  };
  if (token !== undefined) h.Authorization = `Bearer ${token}`;
  return h;
};

describe('HTTP server', () => {
  let server: RunningHttpServer;
  let base: string;

  beforeEach(async () => {
    vi.mocked(createMcpServer).mockClear();
    server = await startHttpServer({ host: '127.0.0.1', port: 0, token: TOKEN });
    base = `http://127.0.0.1:${server.port}`;
  });

  afterEach(async () => {
    await server.close();
  });

  it('GET /healthz returns 200 without authentication', async () => {
    const res = await fetch(`${base}/healthz`);
    expect(res.status).toBe(200);
    expect(await res.text()).toBe('');
  });

  it('POST /mcp without Authorization returns 401', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(undefined),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
    expect(res.headers.get('www-authenticate')).toMatch(/^Bearer /);
  });

  it('POST /mcp with wrong bearer token returns 401', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders('not-the-token'),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
  });

  it('POST /mcp with wrong token of identical length still returns 401', async () => {
    const sameLength = 'a'.repeat(TOKEN.length);
    expect(sameLength.length).toBe(TOKEN.length);
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(sameLength),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
  });

  it('POST /mcp with valid token initializes and returns the server info', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(200);
    const body = (await parseMcpResponseBody(res)) as {
      jsonrpc: string;
      id: number;
      result: { serverInfo: { name: string } };
    };
    expect(body.jsonrpc).toBe('2.0');
    expect(body.id).toBe(1);
    expect(body.result.serverInfo.name).toBe('onenote-mcp');
  });

  it('POST /mcp with valid token + tools/list returns the registered OneNote tools', async () => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/list',
        params: {},
      }),
    });
    expect(res.status).toBe(200);
    const body = (await parseMcpResponseBody(res)) as {
      result: { tools: { name: string }[] };
    };
    const names = body.result.tools.map((t) => t.name);
    expect(names).toContain('list_notebooks');
    expect(names).toContain('create_page');
    expect(names).toContain('delete_page');
  });

  it('unknown paths return 404', async () => {
    const res = await fetch(`${base}/nope`, {
      headers: { Authorization: `Bearer ${TOKEN}` },
    });
    expect(res.status).toBe(404);
  });

  it('unknown paths return a JSON not_found body with a matching Content-Length', async () => {
    const res = await fetch(`${base}/nope`);
    const text = await res.text();
    expect(res.headers.get('content-type')).toBe('application/json');
    expect(res.headers.get('content-length')).toBe(Buffer.byteLength(text).toString());
    expect(JSON.parse(text)).toEqual({ error: 'not_found' });
  });

  it('GET /healthz ignores the query string', async () => {
    const res = await fetch(`${base}/healthz?probe=liveness`);
    expect(res.status).toBe(200);
  });

  it('non-GET /healthz is not treated as a health check', async () => {
    const res = await fetch(`${base}/healthz`, { method: 'POST' });
    expect(res.status).toBe(404);
  });

  it.each([
    ['a trailing slash', '/mcp/'],
    ['a different case', '/MCP'],
    ['a sub-path', '/mcp/extra'],
  ])('/mcp with %s is not routed to the MCP handler', async (_label, path) => {
    const res = await fetch(`${base}${path}`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(404);
    expect(createMcpServer).not.toHaveBeenCalled();
  });

  it('a query string on /mcp does not bypass the auth gate', async () => {
    const res = await fetch(`${base}/mcp?token=${TOKEN}`, {
      method: 'POST',
      headers: mcpHeaders(undefined),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
    expect(await res.json()).toEqual({ error: 'unauthorized' });
    expect(createMcpServer).not.toHaveBeenCalled();
  });

  it('a query string on /mcp is still served when the token is valid', async () => {
    const res = await fetch(`${base}/mcp?session=1`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(200);
  });

  it('the auth gate applies to every method on /mcp', async () => {
    for (const method of ['GET', 'DELETE', 'PUT']) {
      const res = await fetch(`${base}/mcp`, { method });
      expect(res.status).toBe(401);
    }
    expect(createMcpServer).not.toHaveBeenCalled();
  });

  it.each([
    ['the bare token without a scheme', TOKEN],
    ['a lower-case scheme', `bearer ${TOKEN}`],
    ['a different scheme', `Basic ${TOKEN}`],
    ['a token prefix', `Bearer ${TOKEN.slice(0, -1)}`],
    ['a token with a trailing extra character', `Bearer ${TOKEN}x`],
    ['an empty bearer token', 'Bearer '],
    // Same string length as the real header, but é is two bytes once UTF-8 encoded.
    ['a same-length token with a multi-byte character', `Bearer ${TOKEN.slice(0, -1)}é`],
  ])('POST /mcp with %s returns 401', async (_label, authorization) => {
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: { ...mcpHeaders(undefined), Authorization: authorization },
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(401);
    expect(createMcpServer).not.toHaveBeenCalled();
  });

  it('builds a fresh MCP server for every authorised request', async () => {
    for (let i = 0; i < 2; i += 1) {
      const res = await fetch(`${base}/mcp`, {
        method: 'POST',
        headers: mcpHeaders(TOKEN),
        body: JSON.stringify(initializePayload),
      });
      expect(res.status).toBe(200);
      await res.text();
    }
    expect(createMcpServer).toHaveBeenCalledTimes(2);
  });

  it('returns a 500 internal_error when the MCP server cannot be built', async () => {
    vi.mocked(createMcpServer).mockImplementationOnce(() => {
      throw new Error('tool registration exploded');
    });
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(500);
    expect(res.headers.get('content-type')).toBe('application/json');
    expect(await res.json()).toEqual({
      error: 'internal_error',
      message: 'tool registration exploded',
    });
  });

  it('stringifies non-Error failures in the 500 body', async () => {
    vi.mocked(createMcpServer).mockImplementationOnce(() => {
      throw 'plain string failure';
    });
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(500);
    expect(await res.json()).toEqual({
      error: 'internal_error',
      message: 'plain string failure',
    });
  });

  it('keeps serving requests after one of them failed', async () => {
    vi.mocked(createMcpServer).mockImplementationOnce(() => {
      throw new Error('transient');
    });
    const failed = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(failed.status).toBe(500);
    await failed.text();

    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
    });
    expect(res.status).toBe(200);
    const body = (await parseMcpResponseBody(res)) as { result: { serverInfo: { name: string } } };
    expect(body.result.serverInfo.name).toBe('onenote-mcp');
  });

  it('ends the response instead of hanging when a failure happens after headers were sent', async () => {
    const actual = await vi.importActual<typeof McpIndex>('@/index.js');
    vi.mocked(createMcpServer).mockImplementationOnce(() => {
      const mcp = actual.createServer();
      mcp.close = async () => {
        throw new Error('close failed');
      };
      return mcp;
    });
    const res = await fetch(`${base}/mcp`, {
      method: 'POST',
      headers: mcpHeaders(TOKEN),
      body: JSON.stringify(initializePayload),
      signal: AbortSignal.timeout(3000),
    });
    expect(res.status).toBe(200);
    const body = (await parseMcpResponseBody(res)) as { result: { serverInfo: { name: string } } };
    expect(body.result.serverInfo.name).toBe('onenote-mcp');
  });
});

describe('HTTP server lifecycle', () => {
  it('createHttpServer refuses an empty bearer token', () => {
    expect(() => createHttpServer({ host: '127.0.0.1', port: 0, token: '' })).toThrow(
      /non-empty bearer token/,
    );
  });

  it('startHttpServer reports the bound port and MCP URL when asked for port 0', async () => {
    const running = await startHttpServer({ host: '127.0.0.1', port: 0, token: TOKEN });
    try {
      expect(running.port).toBeGreaterThan(0);
      expect(running.url).toBe(`http://127.0.0.1:${running.port}/mcp`);
    } finally {
      await running.close();
    }
  });

  it('startHttpServer rejects when the port is already in use', async () => {
    const first = await startHttpServer({ host: '127.0.0.1', port: 0, token: TOKEN });
    try {
      await expect(
        startHttpServer({ host: '127.0.0.1', port: first.port, token: TOKEN }),
      ).rejects.toMatchObject({ code: 'EADDRINUSE' });
    } finally {
      await first.close();
    }
  });

  it('close stops the listener and rejects when called twice', async () => {
    const running = await startHttpServer({ host: '127.0.0.1', port: 0, token: TOKEN });
    await running.close();
    await expect(fetch(`http://127.0.0.1:${running.port}/healthz`)).rejects.toThrow();
    await expect(running.close()).rejects.toMatchObject({ code: 'ERR_SERVER_NOT_RUNNING' });
  });
});

describe('config getters for HTTP transport', () => {
  beforeEach(() => {
    delete process.env[HTTP_TOKEN_ENV];
    delete process.env[HTTP_HOST_ENV];
    delete process.env[HTTP_PORT_ENV];
  });

  it('getHttpToken throws when ONENOTE_MCP_HTTP_TOKEN is unset', () => {
    expect(() => getHttpToken()).toThrow(/ONENOTE_MCP_HTTP_TOKEN/);
  });

  it('getHttpToken throws when ONENOTE_MCP_HTTP_TOKEN is whitespace', () => {
    process.env[HTTP_TOKEN_ENV] = '   ';
    expect(() => getHttpToken()).toThrow(/ONENOTE_MCP_HTTP_TOKEN/);
  });

  it('getHttpToken returns trimmed value', () => {
    process.env[HTTP_TOKEN_ENV] = '  secret  ';
    expect(getHttpToken()).toBe('secret');
  });

  it('getHttpHost prefers the override over env and default', () => {
    process.env[HTTP_HOST_ENV] = '10.0.0.1';
    expect(getHttpHost('192.168.1.1')).toBe('192.168.1.1');
    expect(getHttpHost()).toBe('10.0.0.1');
    delete process.env[HTTP_HOST_ENV];
    expect(getHttpHost()).toBe('127.0.0.1');
  });

  it('getHttpPort prefers the override, accepts string env, validates range', () => {
    process.env[HTTP_PORT_ENV] = '4000';
    expect(getHttpPort(8080)).toBe(8080);
    expect(getHttpPort()).toBe(4000);
    expect(getHttpPort('5000')).toBe(5000);
    expect(() => getHttpPort('99999')).toThrow(/Invalid HTTP port/);
    expect(() => getHttpPort('not-a-port')).toThrow(/Invalid HTTP port/);
    delete process.env[HTTP_PORT_ENV];
    expect(getHttpPort()).toBe(3000);
  });
});
