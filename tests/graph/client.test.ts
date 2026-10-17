import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { GraphError, graphRequest, paginate } from '@/graph/client.js';
import { createPage, searchPages, updatePage, getPage, getPageContent, deletePage } from '@/graph/pages.js';
import { createNotebook, listNotebooks } from '@/graph/notebooks.js';
import { createSection, listSections } from '@/graph/sections.js';
import { createSectionGroup, listSectionGroups } from '@/graph/sectionGroups.js';

vi.mock('../../src/auth/index.js', () => ({
  getAccessToken: vi.fn(async () => 'fake-access-token'),
}));

const sleepMock = vi.hoisted(() => vi.fn(async (_ms: number) => undefined));
vi.mock('node:timers/promises', () => ({ setTimeout: sleepMock }));

const fetchMock = vi.fn();
const originalFetch = globalThis.fetch;

const jsonResponse = (body: unknown, init: ResponseInit = {}): Response =>
  new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    ...init,
  });

beforeEach(() => {
  fetchMock.mockReset();
  sleepMock.mockClear();
  globalThis.fetch = fetchMock as typeof fetch;
});

afterEach(() => {
  globalThis.fetch = originalFetch;
  vi.restoreAllMocks();
});

describe('graphRequest', () => {
  it('sends Authorization header and parses JSON', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [{ id: 'a' }] }));
    const result = await graphRequest<{ value: { id: string }[] }>('/me/onenote/notebooks');
    expect(result.value[0]?.id).toBe('a');

    const [url, init] = fetchMock.mock.calls[0]!;
    expect(String(url)).toBe('https://graph.microsoft.com/v1.0/me/onenote/notebooks');
    expect((init as RequestInit).method).toBe('GET');
    expect((init as RequestInit).headers).toMatchObject({
      Authorization: 'Bearer fake-access-token',
      Accept: 'application/json',
    });
  });

  it('appends query params and skips undefined values', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({}));
    await graphRequest('/x', { query: { $top: 5, foo: undefined, $search: '"hi"' } });
    const [url] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.searchParams.get('$top')).toBe('5');
    expect(parsed.searchParams.has('foo')).toBe(false);
    expect(parsed.searchParams.get('$search')).toBe('"hi"');
  });

  it('returns text when parse=text', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response('<html>hi</html>', {
        status: 200,
        headers: { 'Content-Type': 'text/html' },
      }),
    );
    const html = await graphRequest<string>('/me/onenote/pages/p1/content', {
      accept: 'text/html',
      parse: 'text',
    });
    expect(html).toBe('<html>hi</html>');
  });

  it('returns undefined for an empty 200 body instead of throwing on JSON.parse', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response('', { status: 200, headers: { 'Content-Type': 'application/json' } }),
    );
    const result = await graphRequest('/me/onenote/whatever');
    expect(result).toBeUndefined();
  });

  it('returns undefined for 204 No Content', async () => {
    fetchMock.mockResolvedValueOnce(new Response(null, { status: 204 }));
    const result = await graphRequest('/me/onenote/pages/p1', {
      method: 'DELETE',
      parse: 'none',
    });
    expect(result).toBeUndefined();
  });

  it('retries on 429 then succeeds', async () => {
    fetchMock
      .mockResolvedValueOnce(
        new Response('{"error":{"code":"TooManyRequests","message":"slow down"}}', {
          status: 429,
          headers: { 'Retry-After': '0', 'Content-Type': 'application/json' },
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ ok: true }));

    const result = await graphRequest<{ ok: boolean }>('/x');
    expect(result.ok).toBe(true);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('throws a shaped GraphError on a non-retryable error', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(
        JSON.stringify({ error: { code: 'NotFound', message: 'no such page' } }),
        { status: 404, headers: { 'Content-Type': 'application/json' } },
      ),
    );
    await expect(graphRequest('/me/onenote/pages/missing')).rejects.toMatchObject({
      name: 'GraphError',
      status: 404,
      code: 'NotFound',
      message: 'no such page',
    });
  });

  it('GraphError survives non-JSON error bodies', async () => {
    fetchMock.mockResolvedValueOnce(new Response('plain text error', { status: 400 }));
    const err = await graphRequest('/x').catch((e) => e as GraphError);
    expect(err).toBeInstanceOf(GraphError);
    expect(err.status).toBe(400);
    expect(err.message).toContain('plain text error');
  });
});

describe('graphRequest retries and error shaping', () => {
  const errorResponse = (status: number, headers: Record<string, string> = {}): Response =>
    new Response(JSON.stringify({ error: { code: `E${status}`, message: `status ${status}` } }), {
      status,
      headers: { 'Content-Type': 'application/json', ...headers },
    });

  it.each([429, 502, 503, 504])('retries a %i response', async (status) => {
    fetchMock
      .mockResolvedValueOnce(errorResponse(status, { 'Retry-After': '0' }))
      .mockResolvedValueOnce(jsonResponse({ ok: true }));

    await expect(graphRequest('/x')).resolves.toEqual({ ok: true });
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it.each([400, 401, 403, 409, 500])('does not retry a %i response', async (status) => {
    fetchMock.mockResolvedValueOnce(errorResponse(status));

    await expect(graphRequest('/x')).rejects.toMatchObject({ status, code: `E${status}` });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(sleepMock).not.toHaveBeenCalled();
  });

  it('gives up after the initial attempt plus four retries and throws the last error', async () => {
    for (let i = 0; i < 5; i += 1) {
      fetchMock.mockResolvedValueOnce(
        new Response(JSON.stringify({ error: { code: 'Throttled', message: `attempt ${i}` } }), {
          status: 429,
          headers: { 'Retry-After': '0' },
        }),
      );
    }

    await expect(graphRequest('/x')).rejects.toMatchObject({
      status: 429,
      code: 'Throttled',
      message: 'attempt 4',
    });
    expect(fetchMock).toHaveBeenCalledTimes(5);
    expect(sleepMock).toHaveBeenCalledTimes(4);
  });

  it('honours Retry-After in seconds', async () => {
    fetchMock
      .mockResolvedValueOnce(errorResponse(503, { 'Retry-After': '3' }))
      .mockResolvedValueOnce(jsonResponse({}));

    await graphRequest('/x');
    expect(sleepMock).toHaveBeenCalledWith(3000);
  });

  it('caps Retry-After at 30 seconds', async () => {
    fetchMock
      .mockResolvedValueOnce(errorResponse(429, { 'Retry-After': '3600' }))
      .mockResolvedValueOnce(jsonResponse({}));

    await graphRequest('/x');
    expect(sleepMock).toHaveBeenCalledWith(30_000);
  });

  it.each([
    ['missing', undefined],
    ['an HTTP date', 'Wed, 21 Oct 2026 07:28:00 GMT'],
    ['negative', '-5'],
  ])('falls back to jittered exponential backoff when Retry-After is %s', async (_label, value) => {
    vi.spyOn(Math, 'random').mockReturnValue(0.5);
    const headers: Record<string, string> = value === undefined ? {} : { 'Retry-After': value };
    for (let i = 0; i < 4; i += 1) fetchMock.mockResolvedValueOnce(errorResponse(503, headers));
    fetchMock.mockResolvedValueOnce(jsonResponse({}));

    await graphRequest('/x');
    expect(sleepMock.mock.calls.map(([ms]) => ms)).toEqual([250, 500, 1000, 2000]);
  });

  it('carries the request-id header onto the GraphError', async () => {
    fetchMock.mockResolvedValueOnce(errorResponse(404, { 'request-id': 'req-abc' }));

    await expect(graphRequest('/x')).rejects.toMatchObject({ requestId: 'req-abc' });
  });

  it('falls back to the status line when the error body is empty', async () => {
    fetchMock.mockResolvedValueOnce(new Response('', { status: 403, statusText: 'Forbidden' }));

    const err = await graphRequest('/x').catch((e: unknown) => e as GraphError);
    expect(err.message).toBe('Graph request failed with 403 Forbidden');
    expect(err.code).toBeUndefined();
    expect(err.requestId).toBeUndefined();
  });

  it('keeps the status line message when the error object has a code but no message', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response(JSON.stringify({ error: { code: 'BadThing' } }), {
        status: 400,
        statusText: 'Bad Request',
      }),
    );

    await expect(graphRequest('/x')).rejects.toMatchObject({
      code: 'BadThing',
      message: 'Graph request failed with 400 Bad Request',
    });
  });

  it('uses the raw text when the JSON body has no error envelope', async () => {
    const body = JSON.stringify({ detail: 'something odd' });
    fetchMock.mockResolvedValueOnce(new Response(body, { status: 400 }));

    await expect(graphRequest('/x')).rejects.toMatchObject({ message: body, code: undefined });
  });

  it('truncates long non-JSON error bodies to 500 characters', async () => {
    fetchMock.mockResolvedValueOnce(new Response('x'.repeat(2000), { status: 400 }));

    const err = await graphRequest('/x').catch((e: unknown) => e as GraphError);
    expect(err.message).toBe('x'.repeat(500));
  });

  it('keeps the status line message when the error body cannot be read', async () => {
    const response = new Response('ignored', { status: 400, statusText: 'Bad Request' });
    vi.spyOn(response, 'text').mockRejectedValue(new Error('socket hang up'));
    fetchMock.mockResolvedValueOnce(response);

    await expect(graphRequest('/x')).rejects.toMatchObject({
      status: 400,
      message: 'Graph request failed with 400 Bad Request',
    });
  });
});

describe('searchPages', () => {
  it('slices results down to the caller-supplied limit', async () => {
    // Graph's $top is a page size, not a total cap, so paginate may return more
    // than `limit` rows. searchPages should still respect the caller's limit.
    const makePage = (id: string) => ({ id, title: `t${id}` });
    fetchMock
      .mockResolvedValueOnce(
        jsonResponse({
          value: ['1', '2', '3'].map(makePage),
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?$skiptoken=a',
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: ['4', '5'].map(makePage) }));

    const results = await searchPages({ query: 'hello', limit: 2 });
    expect(results.map((r) => r.id)).toEqual(['1', '2']);
  });
});

describe('updatePage', () => {
  it('PATCHes /content with a JSON command array', async () => {
    fetchMock.mockResolvedValueOnce(new Response(null, { status: 204 }));

    await updatePage('page-123', [
      { target: 'body', action: 'append', content: '<p>hi</p>' },
      { target: '#abc', action: 'delete' },
    ]);

    const [url, init] = fetchMock.mock.calls[0]!;
    expect(String(url)).toBe(
      'https://graph.microsoft.com/v1.0/me/onenote/pages/page-123/content',
    );
    expect((init as RequestInit).method).toBe('PATCH');
    expect((init as RequestInit).headers).toMatchObject({
      'Content-Type': 'application/json',
    });
    // Targets are sent without a leading `#`. The caller's `#abc` becomes `abc`.
    expect(JSON.parse((init as RequestInit).body as string)).toEqual([
      { target: 'body', action: 'append', content: '<p>hi</p>' },
      { target: 'abc', action: 'delete' },
    ]);
  });
});

describe('createNotebook', () => {
  it('POSTs displayName to /me/onenote/notebooks with $select shaping the response', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 'n1', displayName: 'My Book' }));
    const notebook = await createNotebook('My Book');
    expect(notebook.id).toBe('n1');

    const [url, init] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/notebooks');
    expect(parsed.searchParams.get('$select')).toContain('isDefault');
    expect((init as RequestInit).method).toBe('POST');
    expect(JSON.parse((init as RequestInit).body as string)).toEqual({
      displayName: 'My Book',
    });
  });
});

describe('createSection', () => {
  it('POSTs to the notebook sections endpoint when notebookId is given, with $select and $expand', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 's1', displayName: 'Work' }));
    const section = await createSection({ notebookId: 'nb-abc' }, 'Work');
    expect(section.id).toBe('s1');

    const [url, init] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/notebooks/nb-abc/sections');
    expect(parsed.searchParams.get('$select')).toContain('parentNotebook');
    expect(parsed.searchParams.get('$expand')).toContain('parentNotebook');
    expect(parsed.searchParams.get('$expand')).toContain('parentSectionGroup');
    expect((init as RequestInit).method).toBe('POST');
    expect(JSON.parse((init as RequestInit).body as string)).toEqual({ displayName: 'Work' });
  });

  it('POSTs to the sectionGroups sections endpoint when sectionGroupId is given', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 's2', displayName: 'Sub' }));
    await createSection({ sectionGroupId: 'sg-xyz' }, 'Sub');
    const [url] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/sectionGroups/sg-xyz/sections');
  });
});

describe('createPage', () => {
  it('sends application/xhtml+xml when no attachments are provided', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 'p1', title: 'Hi' }));
    await createPage({ sectionId: 'sec-1', html: '<html><body>hi</body></html>' });

    const [url, init] = fetchMock.mock.calls[0]!;
    expect(String(url)).toBe(
      'https://graph.microsoft.com/v1.0/me/onenote/sections/sec-1/pages',
    );
    expect((init as RequestInit).method).toBe('POST');
    expect((init as RequestInit).headers).toMatchObject({
      'Content-Type': 'application/xhtml+xml',
    });
    expect((init as RequestInit).body).toBe('<html><body>hi</body></html>');
  });

  it('accepts a path-only attachment object as long as bytes are pre-loaded', async () => {
    // Path resolution lives in the tool layer; the graph wrapper accepts bytes.
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 'p3', title: 'Hi' }));
    await createPage({
      sectionId: 's',
      html: '<html><body><img src="name:x"/></body></html>',
      attachments: [{ name: 'x', contentType: 'image/png', data: new Uint8Array([1, 2]) }],
    });
    expect((fetchMock.mock.calls[0]![1] as RequestInit).body).toBeInstanceOf(FormData);
  });

  it('builds a multipart/form-data body when attachments are present', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 'p2', title: 'With image' }));
    await createPage({
      sectionId: 'sec-2',
      html: '<html><body><img src="name:img1"/></body></html>',
      attachments: [
        {
          name: 'img1',
          contentType: 'image/png',
          data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
        },
      ],
    });

    const [, init] = fetchMock.mock.calls[0]!;
    // Content-Type is omitted so fetch + FormData injects the multipart boundary.
    const headers = (init as RequestInit).headers as Record<string, string>;
    expect(headers['Content-Type']).toBeUndefined();
    expect((init as RequestInit).body).toBeInstanceOf(FormData);
    const form = (init as RequestInit).body as FormData;
    expect(form.get('Presentation')).toBeInstanceOf(Blob);
    expect(form.get('img1')).toBeInstanceOf(Blob);
    expect((form.get('img1') as Blob).type).toBe('image/png');
  });
});

describe('section groups', () => {
  it('listSectionGroups hits the notebook-scoped endpoint with a notebook id, expanding both parents', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [] }));
    await listSectionGroups('nb-1');
    const [url] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/notebooks/nb-1/sectionGroups');
    expect(parsed.searchParams.get('$expand')).toContain('parentNotebook');
    expect(parsed.searchParams.get('$expand')).toContain('parentSectionGroup');
  });

  it('listSectionGroups hits the global endpoint without a notebook id', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [] }));
    await listSectionGroups();
    const [url] = fetchMock.mock.calls[0]!;
    expect(String(url)).toContain('/me/onenote/sectionGroups');
    expect(String(url)).not.toContain('/notebooks/');
  });

  it('createSectionGroup nests under a parent section group when given one, with $expand', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 'g1', displayName: 'Inner' }));
    await createSectionGroup({ sectionGroupId: 'sg-1' }, 'Inner');
    const [url, init] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/sectionGroups/sg-1/sectionGroups');
    expect(parsed.searchParams.get('$expand')).toContain('parentSectionGroup');
    expect(JSON.parse((init as RequestInit).body as string)).toEqual({ displayName: 'Inner' });
  });
});
describe('paginate', () => {
  it('returns an empty list when the response body is empty instead of dereferencing undefined', async () => {
    // graphRequest resolves to undefined for an empty 200 body; paginate must
    // not read `.value` off it.
    fetchMock.mockResolvedValueOnce(new Response('', { status: 200 }));
    await expect(paginate('/x')).resolves.toEqual([]);
  });

  it('returns an empty list for a 204 response', async () => {
    fetchMock.mockResolvedValueOnce(new Response(null, { status: 204 }));
    await expect(paginate('/x')).resolves.toEqual([]);
  });

  it('stops when a nextLink page comes back empty', async () => {
    fetchMock
      .mockResolvedValueOnce(
        jsonResponse({
          value: [{ id: '1' }],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?$skiptoken=abc',
        }),
      )
      .mockResolvedValueOnce(new Response('', { status: 200 }));

    const results = await paginate<{ id: string }>('/x');
    expect(results.map((r) => r.id)).toEqual(['1']);
  });

  it('follows @odata.nextLink until exhausted', async () => {
    fetchMock
      .mockResolvedValueOnce(
        jsonResponse({
          value: [{ id: '1' }, { id: '2' }],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/x?$skiptoken=abc',
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ value: [{ id: '3' }] }));

    const results = await paginate<{ id: string }>('/x');
    expect(results.map((r) => r.id)).toEqual(['1', '2', '3']);
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });
});

describe('listNotebooks', () => {
  it('GETs /me/onenote/notebooks with $select and $orderby', async () => {
    fetchMock.mockResolvedValueOnce(
      jsonResponse({ value: [{ id: 'nb-1', displayName: 'Work' }] }),
    );
    const notebooks = await listNotebooks();
    expect(notebooks).toEqual([{ id: 'nb-1', displayName: 'Work' }]);

    const [url] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/notebooks');
    expect(parsed.searchParams.get('$select')).toContain('displayName');
    expect(parsed.searchParams.get('$orderby')).toBe('displayName');
  });
});

describe('listSections', () => {
  it('GETs all sections when no notebookId is given', async () => {
    fetchMock.mockResolvedValueOnce(
      jsonResponse({ value: [{ id: 's1', displayName: 'General' }] }),
    );
    const sections = await listSections();
    expect(sections).toEqual([{ id: 's1', displayName: 'General' }]);

    const [url] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/sections');
    expect(parsed.searchParams.get('$expand')).toContain('parentNotebook');
  });

  it('scopes to a notebook when notebookId is given', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [] }));
    await listSections('nb-abc');

    const [url] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/notebooks/nb-abc/sections');
  });

  it('encodes special characters in notebookId', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ value: [] }));
    await listSections('id with spaces');

    const [url] = fetchMock.mock.calls[0]!;
    expect(String(url)).toContain('id%20with%20spaces');
  });
});

describe('getPage', () => {
  it('GETs a single page by ID with $select and $expand', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 'p1', title: 'Hello' }));
    const page = await getPage('p1');
    expect(page).toEqual({ id: 'p1', title: 'Hello' });

    const [url] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/pages/p1');
    expect(parsed.searchParams.get('$select')).toContain('title');
    expect(parsed.searchParams.get('$expand')).toContain('parentSection');
  });

  it('encodes special characters in pageId', async () => {
    fetchMock.mockResolvedValueOnce(jsonResponse({ id: 'a/b' }));
    await getPage('a/b');

    const [url] = fetchMock.mock.calls[0]!;
    expect(String(url)).toContain('a%2Fb');
  });
});

describe('getPageContent', () => {
  it('GETs page content as text/html', async () => {
    fetchMock.mockResolvedValueOnce(
      new Response('<html><body>content</body></html>', {
        status: 200,
        headers: { 'Content-Type': 'text/html' },
      }),
    );
    const html = await getPageContent('page-42');
    expect(html).toBe('<html><body>content</body></html>');

    const [url, init] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/pages/page-42/content');
    expect((init as RequestInit).headers).toMatchObject({ Accept: 'text/html' });
  });
});

describe('deletePage', () => {
  it('DELETEs the page and returns void', async () => {
    fetchMock.mockResolvedValueOnce(new Response(null, { status: 204 }));
    const result = await deletePage('page-99');
    expect(result).toBeUndefined();

    const [url, init] = fetchMock.mock.calls[0]!;
    const parsed = new URL(String(url));
    expect(parsed.pathname).toBe('/v1.0/me/onenote/pages/page-99');
    expect((init as RequestInit).method).toBe('DELETE');
  });

  it('encodes special characters in pageId', async () => {
    fetchMock.mockResolvedValueOnce(new Response(null, { status: 204 }));
    await deletePage('id/with+special');

    const [url] = fetchMock.mock.calls[0]!;
    expect(String(url)).toContain('id%2Fwith%2Bspecial');
  });
});
