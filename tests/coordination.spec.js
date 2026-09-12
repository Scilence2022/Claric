const http = require('http');
const { CoordinationStore } = require('../src/lib/coordination/store.cjs');
const { createCoordinationServer } = require('../scripts/coordination-server.cjs');

function request(port, method, path, body, token = 'secret') {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const req = http.request({ hostname: '127.0.0.1', port, method, path, headers: { ...(data ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) } : {}), authorization: `Bearer ${token}` } }, (res) => {
      const chunks = []; res.on('data', (chunk) => chunks.push(chunk)); res.on('end', () => resolve({ status: res.statusCode, body: JSON.parse(Buffer.concat(chunks).toString()) }));
    });
    req.on('error', reject); if (data) req.write(data); req.end();
  });
}

function requestWithAuthorization(port, method, path, token) {
  return request(port, method, path, undefined, token);
}

describe('coordination store', () => {
  test('isolates rooms and assigns monotonic event sequences', () => {
    const store = new CoordinationStore({ maxEvents: 2 });
    expect(store.append({ version: 1, workspaceId: 'one', documentId: 'doc-a', type: 'metadata', clientId: 'a', payload: { title: 'A' } }).sequence).toBe(1);
    expect(store.append({ version: 1, workspaceId: 'one', documentId: 'doc-a', type: 'presence', clientId: 'a', payload: { active: true } }).sequence).toBe(2);
    expect(store.append({ version: 1, workspaceId: 'one', documentId: 'doc-b', type: 'metadata', clientId: 'b', payload: { mode: 'review' } }).sequence).toBe(3);
    expect(store.snapshot('one')).toMatchObject({ sequence: 3, metadataByDocument: { 'doc-a': { title: 'A' }, 'doc-b': { mode: 'review' } }, presence: { a: { active: true } } });
    expect(store.snapshot('one').events.map((event) => event.sequence)).toEqual([2, 3]);
    expect(store.snapshot('two')).toMatchObject({ sequence: 0, events: [], metadataByDocument: {}, presence: {} });
  });

  test('rejects cross-room and invalid events', () => {
    const store = new CoordinationStore();
    expect(() => store.append({ version: 1, workspaceId: 'one', type: 'unknown', clientId: 'a', payload: {} })).toThrow('Invalid event type');
    expect(() => store.append({ version: 1, workspaceId: 'one', type: 'metadata', clientId: 'a', payload: {}, extra: 'x' })).toThrow('Unknown event field');
    expect(() => store.append({ version: 1, workspaceId: 'other', type: 'metadata', clientId: 'a', payload: {} })).not.toThrow();
  });
});

describe('coordination HTTP server', () => {
  let server; let port;
  beforeAll(async () => { server = createCoordinationServer({ store: new CoordinationStore(), token: 'secret' }); await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve)); port = server.address().port; });
  afterAll(async () => new Promise((resolve) => server.close(resolve)));

  test('supports health, event relay, and workspace snapshot', async () => {
    expect((await request(port, 'GET', '/coordination/healthz')).body.status).toBe('ok');
    expect((await request(port, 'POST', '/coordination/events', { version: 1, workspaceId: 'room-a', documentId: 'doc-a', type: 'presence', clientId: 'pane-1', payload: { active: true } })).status).toBe(201);
    const snapshot = await request(port, 'GET', '/coordination/snapshot?workspaceId=room-a');
    expect(snapshot.body).toMatchObject({ workspaceId: 'room-a', sequence: 1, presence: { 'pane-1': { active: true, documentId: 'doc-a' } } });
  });

  test('enforces bearer auth and validates workspace input', async () => {
    expect((await request(port, 'GET', '/coordination/healthz', null, 'wrong')).status).toBe(401);
    expect((await request(port, 'GET', '/coordination/snapshot?workspaceId=bad%20id')).status).toBe(400);
    expect((await request(port, 'GET', '/coordination/snapshot')).status).toBe(400);
    expect((await request(port, 'GET', '/coordination/snapshot?workspaceId=a&workspaceId=b')).status).toBe(400);
  });

  test('registers two instances and routes a v2 request through issued credentials', async () => {
    const register = async (documentId) => request(port, 'POST', '/coordination/v2/instances/register', {
      workspaceId: 'workspace-v2', documentId,
    });
    const a = await register('document-a');
    const b = await register('document-b');
    expect(a.status).toBe(201);
    expect(a.body.identity).toMatchObject({ workspaceId: 'workspace-v2', documentId: 'document-a' });
    expect(a.body.credential).not.toBe(b.body.credential);

    const envelope = (source, target, type, payload, key) => ({
      version: 2,
      workspaceId: 'workspace-v2',
      source,
      target,
      type,
      correlationId: `correlation-${key}`,
      idempotencyKey: `idempotency-${key}`,
      ttlMs: 30000,
      createdAt: Date.now(),
      payload,
    });
    const requestEvent = envelope(a.body.identity, b.body.identity, 'context.request', {
      requestId: 'request-http-1', scope: 'selection',
    }, 'request-http-1');
    const sent = await request(port, 'POST', '/coordination/v2/envelopes', requestEvent, a.body.credential);
    expect(sent.status).toBe(201);
    const snapshot = await requestWithAuthorization(port, 'GET', '/coordination/v2/snapshot', b.body.credential);
    expect(snapshot.status).toBe(200);
    expect(snapshot.body.documents).toHaveLength(0);
    const announced = await request(port, 'POST', '/coordination/v2/envelopes', envelope(
      b.body.identity, b.body.identity, 'document.announce', { title: 'B' }, 'announce-http-1',
    ), b.body.credential);
    expect(announced.status).toBe(201);
    const events = await requestWithAuthorization(port, 'GET', `/coordination/v2/events?after=0&epoch=${encodeURIComponent(snapshot.body.epoch)}`, b.body.credential);
    expect(events.status).toBe(200);
    expect(events.body.events.some((event) => event.type === 'context.request' && event.target.instanceId === b.body.identity.instanceId)).toBe(true);
  });

  test('rejects malformed JSON and wrong content type without mutating the room', async () => {
    const send = (headers, body) => new Promise((resolve, reject) => {
      const req = http.request({ hostname: '127.0.0.1', port, method: 'POST', path: '/coordination/events', headers: { authorization: 'Bearer secret', ...headers } }, (res) => {
        const chunks = []; res.on('data', (chunk) => chunks.push(chunk)); res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString() }));
      });
      req.on('error', reject); req.end(body);
    });
    expect((await send({ 'content-type': 'text/plain' }, '{}')).status).toBe(415);
    expect((await send({ 'content-type': 'application/json' }, '{')).status).toBe(400);
    expect((await request(port, 'GET', '/coordination/snapshot?workspaceId=room-a')).body.sequence).toBe(1);
  });

  test('returns snapshots that cannot mutate the store', async () => {
    const first = await request(port, 'GET', '/coordination/snapshot?workspaceId=room-a');
    first.body.presence['pane-1'].active = false;
    first.body.events[0].payload.active = false;
    const second = await request(port, 'GET', '/coordination/snapshot?workspaceId=room-a');
    expect(second.body.presence['pane-1'].active).toBe(true);
    expect(second.body.events[0].payload.active).toBe(true);
  });
});
