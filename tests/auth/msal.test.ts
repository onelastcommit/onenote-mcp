import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { LogLevel, type Configuration } from '@azure/msal-node';
import type * as MsalNode from '@azure/msal-node';
import { AUTHORITY, CLIENT_ID_ENV, TENANT_ID_ENV } from '@/config.js';

const constructedConfigs = vi.hoisted((): Configuration[] => []);
const fakeCachePlugin = vi.hoisted(() => ({
  beforeCacheAccess: async (): Promise<void> => undefined,
  afterCacheAccess: async (): Promise<void> => undefined,
}));

vi.mock('@azure/msal-node', async (importOriginal) => {
  const actual = await importOriginal<typeof MsalNode>();
  class FakePublicClientApplication {
    constructor(public readonly config: Configuration) {
      constructedConfigs.push(config);
    }
  }
  return { ...actual, PublicClientApplication: FakePublicClientApplication };
});

vi.mock('../../src/auth/tokenCache.js', () => ({
  createFileCachePlugin: () => fakeCachePlugin,
}));

const savedEnv: Record<string, string | undefined> = {};

// getMsalClient memoises at module scope, so each case needs a fresh module.
const loadGetMsalClient = async () => {
  vi.resetModules();
  const mod = await import('@/auth/msal.js');
  return mod.getMsalClient;
};

const loggerCallbackFor = (config: Configuration) => {
  const callback = config.system?.loggerOptions?.loggerCallback;
  if (!callback) throw new Error('loggerCallback not configured');
  return callback;
};

beforeEach(() => {
  constructedConfigs.length = 0;
  for (const key of [CLIENT_ID_ENV, TENANT_ID_ENV]) savedEnv[key] = process.env[key];
  process.env[CLIENT_ID_ENV] = 'client-id-123';
  delete process.env[TENANT_ID_ENV];
});

afterEach(() => {
  for (const [key, value] of Object.entries(savedEnv)) {
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }
  vi.restoreAllMocks();
});

describe('getMsalClient', () => {
  it('builds the client from the configured client ID, default authority and file cache plugin', async () => {
    const getMsalClient = await loadGetMsalClient();
    getMsalClient();

    expect(constructedConfigs).toHaveLength(1);
    const config = constructedConfigs[0]!;
    expect(config.auth).toEqual({ clientId: 'client-id-123', authority: AUTHORITY });
    expect(config.cache?.cachePlugin).toBe(fakeCachePlugin);
    expect(config.system?.loggerOptions?.piiLoggingEnabled).toBe(false);
  });

  it('uses the tenant-specific authority when a tenant ID is set', async () => {
    process.env[TENANT_ID_ENV] = 'contoso-tenant';
    const getMsalClient = await loadGetMsalClient();
    getMsalClient();

    expect(constructedConfigs[0]!.auth.authority).toBe(
      'https://login.microsoftonline.com/contoso-tenant',
    );
  });

  it('returns the same instance on repeat calls without reconstructing', async () => {
    const getMsalClient = await loadGetMsalClient();
    const first = getMsalClient();
    const second = getMsalClient();

    expect(second).toBe(first);
    expect(constructedConfigs).toHaveLength(1);
  });

  it('throws without constructing or caching a client when the client ID is missing', async () => {
    delete process.env[CLIENT_ID_ENV];
    const getMsalClient = await loadGetMsalClient();

    expect(() => getMsalClient()).toThrow(CLIENT_ID_ENV);
    expect(constructedConfigs).toHaveLength(0);

    process.env[CLIENT_ID_ENV] = 'late-client-id';
    getMsalClient();
    expect(constructedConfigs[0]!.auth.clientId).toBe('late-client-id');
  });
});

describe('getMsalClient logger', () => {
  it.each([
    ['Error', LogLevel.Error],
    ['Warning', LogLevel.Warning],
  ])('writes %s messages to stderr with an [msal] prefix', async (_name, level) => {
    const stderr = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
    const stdout = vi.spyOn(process.stdout, 'write').mockImplementation(() => true);
    const getMsalClient = await loadGetMsalClient();
    getMsalClient();

    loggerCallbackFor(constructedConfigs[0]!)(level, 'token refresh failed', false);

    expect(stderr).toHaveBeenCalledWith('[msal] token refresh failed\n');
    // stdout carries the MCP stdio protocol; any MSAL output there would corrupt it.
    expect(stdout).not.toHaveBeenCalled();
  });

  it.each([
    ['Info', LogLevel.Info],
    ['Verbose', LogLevel.Verbose],
    ['Trace', LogLevel.Trace],
  ])('drops %s messages entirely', async (_name, level) => {
    const stderr = vi.spyOn(process.stderr, 'write').mockImplementation(() => true);
    const stdout = vi.spyOn(process.stdout, 'write').mockImplementation(() => true);
    const getMsalClient = await loadGetMsalClient();
    getMsalClient();

    loggerCallbackFor(constructedConfigs[0]!)(level, 'chatty detail', false);

    expect(stderr).not.toHaveBeenCalled();
    expect(stdout).not.toHaveBeenCalled();
  });

  it('asks MSAL itself to filter at Warning level', async () => {
    const getMsalClient = await loadGetMsalClient();
    getMsalClient();

    expect(constructedConfigs[0]!.system?.loggerOptions?.logLevel).toBe(LogLevel.Warning);
  });
});
