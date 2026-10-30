import { createSharedEvidenceStore, createSharedEvidenceReader, validateSharedEvidencePage, SHARED_EVIDENCE_LIMITS } from '../src/taskpane/shared-evidence.js';
import { createContextRequestManager } from '../src/taskpane/context-requests.js';
const { webcrypto } = require('crypto');
const { CoordinationState } = require('../src/lib/coordination/state.cjs');

const identity = { workspaceId: 'w', documentId: 'source', instanceId: 'source-instance' };
const snapshot = (text = 'Observed source evidence', overrides = {}) => ({ text, scope: 'selection', snapshotId: 'snapshot-1',
    documentRevision: 'revision-1', sourceDocumentId: identity.documentId, sourceInstanceId: identity.instanceId, ...overrides });
const create = (options = {}) => createSharedEvidenceStore({ identity, cryptoImpl: webcrypto, ...options });

test('publishes only explicit bounded snapshots and exposes metadata without content', async () => {
    const store = create();
    expect(store.catalog()).toEqual([]);
    const source = snapshot();
    const meta = await store.publish(source, { title: 'Payment terms' });
    source.text = 'mutated';
    expect(meta).toMatchObject({ source: identity, digest: expect.stringMatching(/^[a-f0-9]{64}$/), scope: 'selection' });
    expect(JSON.stringify(store.catalog())).not.toContain('Observed source evidence');
    meta.source.instanceId = 'spoof';
    expect(store.catalog()[0].source).toEqual(identity);
    expect(store.read({ artifactId: meta.artifactId, digest: meta.digest }).text).toBe('Observed source evidence');
    await expect(store.publish(snapshot('x'.repeat(12001)))).rejects.toMatchObject({ code: 'INVALID_EVIDENCE' });
    await expect(store.publish(snapshot(''))).rejects.toMatchObject({ code: 'INVALID_EVIDENCE' });
    await expect(store.publish(snapshot('x', { sourceInstanceId: 'other' }))).rejects.toMatchObject({ code: 'INVALID_EVIDENCE' });
    await expect(store.publish(snapshot(), { ttlMs: 999 })).rejects.toMatchObject({ code: 'INVALID_EVIDENCE' });
});

test('pages and searches without splitting surrogate pairs or accepting invalid versions', async () => {
    const store = create();
    const meta = await store.publish(snapshot('a😀b payment terms', { truncation: { text: true } }));
    const query = { artifactId: meta.artifactId, digest: meta.digest };
    const first = store.read({ ...query, limit: 2 });
    expect(first).toMatchObject({ text: 'a', offset: 0, nextOffset: 1, truncated: true });
    expect(store.read({ ...query, offset: first.nextOffset, limit: 2 }).text).toBe('😀');
    expect(store.read({ ...query, query: 'PAYMENT' }).text).toContain('payment terms');
    expect(store.read({ ...query, query: 'absent' })).toMatchObject({ matched: false, text: '', nextOffset: null });
    expect(() => store.read({ ...query, digest: 'wrong' })).toThrow('fingerprint');
    for (const options of [{ offset: -1 }, { limit: 4000 }, { offset: 999 }, { query: 'x'.repeat(201) }, { offset: 1, limit: 1 }]) {
        expect(() => store.read({ ...query, ...options })).toThrow();
    }
});

test('expires, revokes, bounds capacity and permanently disposes the workspace', async () => {
    let now = 1000;
    const store = create({ clock: () => now });
    const first = await store.publish(snapshot(), { ttlMs: 1000 });
    now += 1000;
    expect(store.catalog()).toEqual([]);
    expect(() => store.metadata(first.artifactId)).toThrow('expired');
    for (let i = 0; i < SHARED_EVIDENCE_LIMITS.records; i++) await store.publish(snapshot('x'));
    await expect(store.publish(snapshot())).rejects.toMatchObject({ code: 'EVIDENCE_CAPACITY' });
    expect(store.remove(store.catalog()[0].artifactId)).toBe(true);
    await store.publish(snapshot());
    store.dispose();
    expect(store.catalog()).toEqual([]);
    await expect(store.publish(snapshot())).rejects.toMatchObject({ code: 'EVIDENCE_UNAVAILABLE' });
});

test('revalidates the source, identity and advertised version before returning evidence', async () => {
    const store = create();
    const meta = await store.publish(snapshot());
    const agent = { getDocumentRevision: jest.fn(async () => 'revision-1'), readContext: jest.fn(async () => 'local') };
    const reader = createSharedEvidenceReader({ store, agent, identity });
    const request = { scope: 'shared_evidence', artifactIds: [meta.artifactId], digest: meta.digest };
    const result = await reader.readContext(request);
    expect(validateSharedEvidencePage(result, meta, identity).text).toBe('Observed source evidence');
    for (const changed of [{ digest: 'wrong' }, { sourceRevision: 'wrong' }, { sourceSnapshotId: 'wrong' }]) {
        expect(() => validateSharedEvidencePage(result, { ...meta, ...changed }, identity)).toThrow('advertised version');
    }
    expect(() => validateSharedEvidencePage({ ...result, evidence: { ...result.evidence, nextOffset: -1 } }, meta, identity)).toThrow();
    expect(await reader.readContext({ scope: 'document' })).toBe('local');
    agent.getDocumentRevision.mockResolvedValue('revision-2');
    await expect(reader.readContext(request)).rejects.toMatchObject({ code: 'EVIDENCE_SOURCE_CHANGED' });
    const controller = new AbortController(); controller.abort();
    await expect(reader.readContext(request, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    await expect(reader.readContext({ ...request, artifactIds: [] })).rejects.toMatchObject({ code: 'INVALID_EVIDENCE_QUERY' });
    store.remove(meta.artifactId);
    await expect(reader.readContext(request)).rejects.toMatchObject({ code: 'EVIDENCE_UNAVAILABLE' });
});

test('shares catalogue metadata across authenticated peers and reads pages through the existing context protocol', async () => {
    const state = new CoordinationState({ clock: () => 1000 });
    const owner = state.register({ workspaceId: 'w', documentId: 'source' }, 'local');
    const peer = state.register({ workspaceId: 'w', documentId: 'target' }, 'local');
    const stranger = state.register({ workspaceId: 'other', documentId: 'outside' }, 'local');
    let sequence = 0;
    const append = (session, target, type, payload, metadata = {}) => state.append({ version: 2, workspaceId: session.identity.workspaceId,
        source: session.identity, target, type, payload, createdAt: 1000, ttlMs: 30000,
        correlationId: `c-${++sequence}`, idempotencyKey: `i-${sequence}`, ...metadata }, session.credential, 'local');
    const store = create({ identity: owner.identity, clock: () => 1000 });
    const meta = await store.publish(snapshot('a'.repeat(3500) + ' payment terms', { sourceInstanceId: owner.identity.instanceId }));
    append(owner, owner.identity, 'artifact.announced', { artifactId: meta.artifactId, shared: true, mediaType: 'text/plain', digest: meta.digest,
        size: meta.size, textLength: meta.textLength, scope: meta.scope, sourceRevision: meta.sourceRevision, sourceSnapshotId: meta.sourceSnapshotId });
    expect(state.snapshot(peer.credential, 'local').artifacts[0].artifactId).toBe(meta.artifactId);
    expect(state.snapshot(stranger.credential, 'local').artifacts).toEqual([]);
    expect(JSON.stringify(state.snapshot(peer.credential, 'local'))).not.toContain('payment terms');
    let outgoing;
    let response;
    const remoteAgent = createSharedEvidenceReader({ store, identity: owner.identity, agent: { getDocumentRevision: async () => 'revision-1' } });
    const sender = createContextRequestManager({ identity: peer.identity, clock: () => 1000, agent: { readContext: async () => ({}) }, client: {
        identity: peer.identity, requestContext: async (target, payload, metadata) => { outgoing = append(peer, target, 'context.request', payload, metadata); }, respondContext: async () => {},
    } });
    const receiver = createContextRequestManager({ identity: owner.identity, clock: () => 1000, agent: remoteAgent, client: {
        identity: owner.identity, requestContext: async () => {}, respondContext: async (event, payload, metadata) => { response = append(owner, event.source, 'context.response', payload, metadata); },
    } });
    const pending = sender.requestContext(owner.identity, { scope: 'shared_evidence', artifactIds: [meta.artifactId], digest: meta.digest, offset: 3400, limit: 500, query: 'payment' });
    await receiver.handleEvent(outgoing);
    await sender.handleEvent(response);
    const result = await pending;
    expect(result.snapshot.text).toContain('payment terms');
    expect(result.snapshot.evidence.offset).toBe(3400);
    append(owner, owner.identity, 'artifact.removed', { artifactId: meta.artifactId }, { correlationId: state.snapshot(peer.credential, 'local').artifacts[0].correlationId });
    expect(state.snapshot(peer.credential, 'local').artifacts).toEqual([]);
    sender.dispose(); receiver.dispose(); store.dispose();
});
