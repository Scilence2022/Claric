const {
  createCoordinationIdentity,
  filterCoordinationSnapshot,
  createCoordinationTransport,
} = require('../src/taskpane/coordination-client.js');

describe('coordination browser client', () => {
  test('creates stable workspace and explicit ephemeral document fallback', () => {
    const storage = { values: {}, getItem(key) { return this.values[key] || null; }, setItem(key, value) { this.values[key] = value; } };
    const first = createCoordinationIdentity({ locationObject: { origin: 'https://localhost:3001' }, storage });
    const second = createCoordinationIdentity({ locationObject: { origin: 'https://localhost:3001' }, storage });
    expect(first.workspaceId).toBe(second.workspaceId);
    expect(first.instanceId).not.toBe(second.instanceId);
    expect(first.documentEphemeral).toBe(true);
  });

  test('filters document-specific event and metadata scope', () => {
    const identity = { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'self' };
    const result = filterCoordinationSnapshot({ workspaceId: 'workspace-a', sequence: 3, presence: { self: { clientId: 'self' }, other: { clientId: 'other', documentId: 'document-b' } }, metadataByDocument: { '*': { shared: true }, 'document-a': { title: 'A' }, 'document-b': { secret: true } }, events: [{ documentId: 'document-a', payload: { ok: true } }, { documentId: 'document-b', payload: { no: true } }] }, identity);
    expect(result.metadata).toEqual({ shared: true, title: 'A' });
    expect(result.events).toHaveLength(1);
    expect(result.events[0].payload).toEqual({ ok: true });
    expect(result.presence.other.documentId).toBe('document-b');
  });

  test('posts, polls, and stops without leaving a timer', async () => {
    const calls = [];
    let callback;
    const fetchImpl = async (url, options = {}) => {
      calls.push({ url, options });
      return { ok: true, status: 200, json: async () => ({ workspaceId: 'workspace-a', sequence: 1, presence: {}, metadataByDocument: {}, events: [] }) };
    };
    const transport = createCoordinationTransport({ fetchImpl, intervalMs: 20, setIntervalImpl: (fn) => { callback = fn; return 1; }, clearIntervalImpl: () => { callback = null; } });
    const identity = { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'instance-a', documentEphemeral: true };
    await transport.start(identity);
    expect(calls[0].options.method).toBe('POST');
    expect(calls.some((call) => call.url.includes('/snapshot?workspaceId=workspace-a'))).toBe(true);
    expect(callback).toBeInstanceOf(Function);
    await transport.stop(identity);
    expect(callback).toBeNull();
    expect(calls[calls.length - 1].options.method).toBe('POST');
    expect(JSON.parse(calls[calls.length - 1].options.body).payload.active).toBe(false);
  });
});
