const {
  createCoordinationIdentity,
  filterCoordinationSnapshot,
  createCoordinationTransport,
  createV2CoordinationTransport,
  createCoordinationClient,
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

  test('registers v2, retains server identity, and polls targeted events with cursor', async () => {
    const calls = [];
    const issued = { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'server-instance' };
    const fetchImpl = async (url, options = {}) => {
      calls.push({ url, options });
      if (url.endsWith('/instances/register')) return { ok: true, status: 201, json: async () => ({ version: 2, credential: 'credential-not-for-logs', identity: issued, expiresAt: 200000 }) };
      if (url.endsWith('/snapshot')) return { ok: true, status: 200, json: async () => ({ version: 2, workspaceId: 'workspace-a', epoch: 'epoch-a', cursor: 4, documents: [] }) };
      return { ok: true, status: 200, json: async () => ({ epoch: 'epoch-a', cursor: 5, events: [{ sequence: 5, source: issued, target: issued, type: 'document.announce', payload: {} }] }) };
    };
    const transport = createV2CoordinationTransport({ fetchImpl, token: 'bootstrap-token', clock: () => 100000, setIntervalImpl: null });
    await transport.start({ workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'local-spoof' });
    expect(transport.identity).toEqual(issued);
    expect(transport.expiresAt).toBe(200000);
    expect(transport.cursor).toBe(4);
    expect(calls[0].url).toBe('/coordination/v2/instances/register');
    expect(JSON.parse(calls[0].options.body)).toEqual({ workspaceId: 'workspace-a', documentId: 'document-a' });
    expect(calls[0].options.headers.Authorization).toBe('Bearer bootstrap-token');
    expect(calls[1].options.headers.Authorization).toBe('Bearer credential-not-for-logs');
    await transport.poll();
    expect(calls[4].url).toBe('/coordination/v2/events?after=4&epoch=epoch-a');
    expect(calls[4].options.headers.Authorization).toBe('Bearer credential-not-for-logs');
    expect(transport.cursor).toBe(5);
    await transport.stop();
    expect(calls.some(({ options }) => JSON.stringify(options).includes('credential-not-for-logs'))).toBe(true);
  });

  test('re-snapshots after an expired v2 cursor and publishes identity-safe envelopes', async () => {
    const calls = [];
    const issued = { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'server-instance' };
    let snapshots = 0;
    const fetchImpl = async (url, options = {}) => {
      calls.push({ url, options });
      if (url.endsWith('/instances/register')) return { ok: true, status: 201, json: async () => ({ credential: 'secret', identity: issued, expiresAt: 200000 }) };
      if (url.endsWith('/snapshot')) return { ok: true, status: 200, json: async () => ({ version: 2, workspaceId: 'workspace-a', epoch: 'epoch-new', cursor: snapshots++ ? 9 : 3, documents: [] }) };
      if (url.includes('/events?')) return { ok: false, status: 410, json: async () => ({ error: 'expired' }) };
      return { ok: true, status: 201, json: async () => ({ event: JSON.parse(options.body) }) };
    };
    const transport = createV2CoordinationTransport({ fetchImpl, clock: () => 100000, setIntervalImpl: null });
    await transport.start({ workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'caller-instance' });
    await expect(transport.publishEnvelope({
      type: 'context.request', target: { ...issued, instanceId: 'target-instance' }, source: { ...issued, instanceId: 'forged' },
      correlationId: 'corr-1', idempotencyKey: 'idem-1', ttlMs: 30000, payload: { requestId: 'request-1', scope: 'document' },
    })).rejects.toThrow('Cannot override');
    const result = await transport.publishEnvelope({
      type: 'context.request', target: { ...issued, instanceId: 'target-instance' },
      correlationId: 'corr-1', idempotencyKey: 'idem-1', ttlMs: 30000, payload: { requestId: 'request-1', scope: 'document' },
    });
    const envelope = JSON.parse(calls[calls.length - 1].options.body);
    expect(envelope.source).toEqual(issued);
    expect(envelope.correlationId).toBe('corr-1');
    expect(envelope.idempotencyKey).toBe('idem-1');
    expect(envelope.ttlMs).toBe(30000);
    expect(result.event).toEqual(envelope);
    await transport.poll();
    expect(transport.epoch).toBe('epoch-new');
    expect(transport.cursor).toBe(9);
    expect(calls.some(({ url }) => url.endsWith('/snapshot'))).toBe(true);
  });

  test('falls back to v1 only when v2 registration is unavailable, not on auth errors', async () => {
    const calls = [];
    const fetchImpl = async (url, options = {}) => {
      calls.push({ url, options });
      if (url.endsWith('/instances/register')) return { ok: false, status: 404, json: async () => ({ error: 'Not found' }) };
      return { ok: true, status: 200, json: async () => ({ workspaceId: 'workspace-a', sequence: 1, presence: {}, metadataByDocument: {}, events: [] }) };
    };
    const client = createCoordinationClient({ fetchImpl, setIntervalImpl: null, identity: { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'instance-a' } });
    await client.start();
    expect(client.transport.version).toBeUndefined();
    expect(calls.some(({ url }) => url.endsWith('/snapshot?workspaceId=workspace-a'))).toBe(true);

    const authClient = createCoordinationClient({
      fetchImpl: async (url) => url.endsWith('/instances/register')
        ? { ok: false, status: 401, json: async () => ({ error: 'Unauthorized' }) }
        : { ok: true, status: 200, json: async () => ({}) },
      setIntervalImpl: null,
      identity: { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'instance-a' },
    });
    await expect(authClient.start()).rejects.toMatchObject({ status: 401 });
    expect(authClient.transport.version).toBe(2);
  });

  test.each([401, 403, 405, 503])('never falls back after registration HTTP %i', async (status) => {
    const fetchImpl = jest.fn(async () => ({ ok: false, status }));
    const client = createCoordinationClient({ fetchImpl, setIntervalImpl: null, identity: { workspaceId: 'w', documentId: 'd', instanceId: 'i' } });
    await expect(client.start()).rejects.toMatchObject({ status });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  test('recovers pending requests on 410 and consumes events after discovery without skipping them', async () => {
    const { CoordinationState } = require('../src/lib/coordination/state.cjs');
    const state = new CoordinationState({ clock: () => 100000, maxEvents: 1 });
    const events = [];
    const snapshots = [];
    const fetchImpl = async (url, options = {}) => {
      const credential = options.headers.Authorization?.slice(7);
      const body = options.body ? JSON.parse(options.body) : null;
      try {
        let result;
        if (url.endsWith('/register')) result = state.register(body, 'binding');
        else if (url.endsWith('/snapshot')) result = state.snapshot(credential, 'binding');
        else if (url.endsWith('/envelopes')) result = { event: state.append(body, credential, 'binding') };
        else { const query = new URL(url, 'http://localhost').searchParams; result = state.events(credential, 'binding', Number(query.get('after')), query.get('epoch')); }
        return { ok: true, json: async () => result };
      } catch (error) { return { ok: false, status: error.status || 400 }; }
    };
    const transport = createV2CoordinationTransport({ fetchImpl, clock: () => 100000, setIntervalImpl: null, onEvents: (batch) => events.push(...batch), onSnapshot: (value) => snapshots.push(value) });
    await transport.start({ workspaceId: 'w', documentId: 'a' });
    const other = state.register({ workspaceId: 'w', documentId: 'b' }, 'binding');
    const envelope = { version: 2, workspaceId: 'w', source: other.identity, target: transport.identity, type: 'context.request', payload: { requestId: 'r', scope: 'selection' }, correlationId: 'c', idempotencyKey: 'r', createdAt: 100000, ttlMs: 30000 };
    state.append(envelope, other.credential, 'binding');
    state.append({ ...envelope, type: 'document.announce', target: other.identity, payload: {}, idempotencyKey: 'announce' }, other.credential, 'binding');
    const first = transport.poll();
    expect(transport.poll()).toBe(first);
    await first;
    expect(events).toEqual(expect.arrayContaining([expect.objectContaining({ type: 'context.request', payload: { requestId: 'r', scope: 'selection' }, recovered: true })]));
    expect(snapshots.at(-1).documents).toEqual(expect.arrayContaining([expect.objectContaining({ documentId: 'b' })]));
    await transport.stop();
  });

  describe('v2 event stream (sse mode)', () => {
    const encoder = new TextEncoder();
    const issued = { workspaceId: 'w', documentId: 'a', instanceId: 'srv-a' };
    const other = { workspaceId: 'w', documentId: 'b', instanceId: 'srv-b' };
    const sseData = (obj, eventName = '') => encoder.encode(`${eventName ? `event: ${eventName}\n` : ''}data: ${JSON.stringify(obj)}\n\n`);
    const flush = () => new Promise((resolve) => setTimeout(resolve, 0));
    const blockingReader = (chunks) => {
      let index = 0;
      return { read: jest.fn(() => (index < chunks.length ? Promise.resolve({ done: false, value: chunks[index++] }) : new Promise(() => {}))) };
    };
    const endingReader = (chunks) => {
      let index = 0;
      return { read: jest.fn(async () => (index < chunks.length ? { done: false, value: chunks[index++] } : { done: true, value: undefined })) };
    };
    let timers;
    let errors;
    const timerImpls = () => ({
      setIntervalImpl: (fn) => { timers.interval = fn; return 1; },
      clearIntervalImpl: () => { timers.interval = null; },
      setTimeoutImpl: (fn, delay) => { timers.timeouts.push({ fn, delay }); return timers.timeouts.length; },
      clearTimeoutImpl: () => {},
    });
    const v2Fetch = ({ snapshot, stream, calls }) => async (url, options = {}) => {
      calls.push({ url, options });
      if (url.endsWith('/instances/register')) return { ok: true, status: 201, json: async () => ({ credential: 'cred', identity: issued, expiresAt: 200000 }) };
      if (url.endsWith('/snapshot')) return { ok: true, status: 200, json: async () => snapshot() };
      if (url.includes('/events/stream')) return stream(options);
      if (url.includes('/events?')) return { ok: true, status: 200, json: async () => ({ epoch: 'ep', cursor: 6, events: [] }) };
      return { ok: true, status: 201, json: async () => ({ event: options.body ? JSON.parse(options.body) : {} }) };
    };
    const sseTransport = (fetchImpl, handlers = {}) => createV2CoordinationTransport({
      fetchImpl, clock: () => 100000, mode: 'sse', requestTimeoutMs: 60000, ...timerImpls(),
      onError: (error) => errors.push(error), ...handlers,
    });
    beforeEach(() => {
      timers = { interval: null, timeouts: [] };
      errors = [];
    });

    test('streams events in real time and ticks discovery only while the stream is live', async () => {
      const events = [];
      const calls = [];
      const fetchImpl = v2Fetch({
        calls,
        snapshot: () => ({ version: 2, workspaceId: 'w', epoch: 'ep', cursor: 4, documents: [] }),
        stream: async () => ({ ok: true, status: 200, body: { getReader: () => blockingReader([
          sseData({ epoch: 'ep', cursor: 5, events: [] }),
          sseData({ epoch: 'ep', cursor: 6, events: [{ sequence: 6, source: other, target: issued, type: 'task.submit', payload: { taskId: 't1' } }] }),
          sseData({ epoch: 'ep', cursor: 6, events: [{ sequence: 6, source: other, target: issued, type: 'task.submit', payload: { taskId: 't1-duplicate' } }] }),
        ]) } }),
      });
      const transport = sseTransport(fetchImpl, { onEvents: (batch) => events.push(...batch) });
      await transport.start({ workspaceId: 'w', documentId: 'a' });
      await flush();
      expect(transport.streaming).toBe(true);
      expect(events).toHaveLength(1);
      expect(events[0]).toMatchObject({ type: 'task.submit', payload: { taskId: 't1' } });
      expect(transport.cursor).toBe(6);
      expect(errors).toHaveLength(0);

      calls.length = 0;
      timers.interval();
      await flush();
      expect(calls.some(({ url }) => url.endsWith('/snapshot'))).toBe(true);
      expect(calls.some(({ url }) => url.endsWith('/envelopes'))).toBe(true);
      expect(calls.some(({ url }) => url.includes('/events?'))).toBe(false);
      await transport.stop();
    });

    test('falls back to full polling when the server has no event stream', async () => {
      const calls = [];
      const fetchImpl = v2Fetch({
        calls,
        snapshot: () => ({ version: 2, workspaceId: 'w', epoch: 'ep', cursor: 4, documents: [] }),
        stream: async () => ({ ok: false, status: 404, json: async () => ({ error: 'Not found' }) }),
      });
      const transport = sseTransport(fetchImpl);
      await transport.start({ workspaceId: 'w', documentId: 'a' });
      await flush();
      expect(transport.streaming).toBe(false);
      expect(errors.some((error) => String(error.message).includes('using polling'))).toBe(true);

      calls.length = 0;
      timers.interval();
      await flush();
      expect(calls.some(({ url }) => url.includes('/events?after=4&epoch=ep'))).toBe(true);
      expect(transport.cursor).toBe(6);
      expect(timers.timeouts.some(({ delay }) => delay !== 60000)).toBe(false);
      await transport.stop();
    });

    test('re-snapshots when the stream sends a resync frame', async () => {
      const calls = [];
      let snapshotCount = 0;
      const fetchImpl = v2Fetch({
        calls,
        snapshot: () => ({ version: 2, workspaceId: 'w', epoch: 'ep', cursor: snapshotCount++ ? 9 : 4, documents: [] }),
        stream: async () => ({ ok: true, status: 200, body: { getReader: () => blockingReader([sseData({ reason: 'cursor-expired' }, 'resync')]) } }),
      });
      const transport = sseTransport(fetchImpl);
      await transport.start({ workspaceId: 'w', documentId: 'a' });
      await flush();
      expect(snapshotCount).toBe(2);
      expect(transport.cursor).toBe(9);
      await transport.stop();
    });

    test('aborts the stream on stop without scheduling a reconnect', async () => {
      const calls = [];
      let streamSignal;
      const fetchImpl = v2Fetch({
        calls,
        snapshot: () => ({ version: 2, workspaceId: 'w', epoch: 'ep', cursor: 4, documents: [] }),
        stream: async (options) => { streamSignal = options.signal; return { ok: true, status: 200, body: { getReader: () => blockingReader([]) } }; },
      });
      const transport = sseTransport(fetchImpl);
      await transport.start({ workspaceId: 'w', documentId: 'a' });
      await flush();
      expect(transport.streaming).toBe(true);
      await transport.stop();
      expect(streamSignal.aborted).toBe(true);
      expect(timers.timeouts.some(({ delay }) => delay !== 60000)).toBe(false);
    });

    test('reconnects with the current cursor when the stream closes', async () => {
      const events = [];
      const calls = [];
      const streamUrls = [];
      let streamCount = 0;
      const fetchImpl = v2Fetch({
        calls,
        snapshot: () => ({ version: 2, workspaceId: 'w', epoch: 'ep', cursor: 4, documents: [] }),
        stream: async () => {
          streamUrls.push(calls[calls.length - 1].url);
          streamCount += 1;
          return streamCount === 1
            ? { ok: true, status: 200, body: { getReader: () => endingReader([sseData({ epoch: 'ep', cursor: 5, events: [{ sequence: 5, source: other, target: issued, type: 'task.submit', payload: { taskId: 't2' } }] })]) } }
            : { ok: true, status: 200, body: { getReader: () => blockingReader([]) } };
        },
      });
      const transport = sseTransport(fetchImpl, { onEvents: (batch) => events.push(...batch) });
      await transport.start({ workspaceId: 'w', documentId: 'a' });
      await flush();
      expect(events).toHaveLength(1);
      expect(transport.streaming).toBe(false);
      const reconnect = timers.timeouts.find(({ delay }) => delay !== 60000);
      expect(reconnect).toBeTruthy();
      reconnect.fn();
      await flush();
      expect(streamUrls[1]).toContain('/events/stream?after=5&epoch=ep');
      expect(transport.streaming).toBe(true);
      await transport.stop();
    });
  });
});
