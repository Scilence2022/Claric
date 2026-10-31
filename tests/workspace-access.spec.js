import { createWorkspaceAccess } from '../src/taskpane/workspace-access.js';
const { normalizeDocumentAccess, allowsDocumentAccess } = require('../src/lib/coordination/access.cjs');
const { CoordinationState } = require('../src/lib/coordination/state.cjs');

const local = { workspaceId: 'w', documentId: 'owner', instanceId: 'owner-1' };
const peer = { workspaceId: 'w', documentId: 'peer', instanceId: 'peer-1' };
test('persists scoped document rules, fails closed on corruption, and never overrides master switches', () => {
    const values = new Map();
    const storage = { getItem: (key) => values.get(key), setItem: (key, value) => values.set(key, value) };
    const policy = createWorkspaceAccess({ identity: local, storage });
    expect(policy.allows(peer, 'readContext')).toBe(true);
    policy.update('peer', 'readContext', false);
    expect(createWorkspaceAccess({ identity: local, storage }).allows({ ...peer, instanceId: 'other-window' }, 'readContext')).toBe(false);
    expect(policy.allows(peer, 'requestEdits')).toBe(true);
    policy.update(null, 'requestEdits', false);
    policy.update('peer', 'requestEdits', true);
    expect(policy.allows(peer, 'requestEdits')).toBe(false);
    expect(policy.allows({ ...peer, workspaceId: 'other' }, 'readContext')).toBe(false);
    expect(createWorkspaceAccess({ identity: { ...local, documentId: 'different' }, storage }).allows(peer, 'readContext')).toBe(true);
    values.set([...values.keys()][0], '{corrupt');
    expect(createWorkspaceAccess({ identity: local, storage }).allows(peer, 'requestEdits')).toBe(false);
    storage.setItem = () => { throw new Error('Quota'); };
    expect(() => policy.update('peer', 'readContext', true)).toThrow('Quota');
    expect(policy.allows(peer, 'readContext')).toBe(false);
});

test('rejects malformed permission announcements and unknown operations', () => {
    for (const value of [{ readContext: 'yes' }, { surprise: true }, { peers: [{ documentId: 'x' }, { documentId: 'x' }] }, { peers: [{ documentId: '../x' }] }]) expect(() => normalizeDocumentAccess(value)).toThrow();
    expect(allowsDocumentAccess({}, 'peer', 'delete')).toBe(false);
    expect(allowsDocumentAccess({ readContext: 'yes' }, 'peer', 'readContext')).toBe(false);
});

test('server denies new reads, late read results and edit requests according to owner policy', () => {
    const state = new CoordinationState({ clock: () => 1000 });
    const owner = state.register({ workspaceId: 'w', documentId: 'owner' }, 'local');
    const source = state.register({ workspaceId: 'w', documentId: 'peer' }, 'local');
    let count = 0;
    const send = (session, target, type, payload, correlationId = `c-${++count}`) => state.append({ version: 2, workspaceId: 'w', source: session.identity, target, type, payload,
        createdAt: 1000, ttlMs: 30000, correlationId, idempotencyKey: `i-${++count}` }, session.credential, 'local');
    send(source, owner.identity, 'context.request', { requestId: 'r1', scope: 'document' }, 'read');
    send(owner, owner.identity, 'document.announce', { access: { peers: [{ documentId: 'peer', readContext: false, requestEdits: false }] } });
    expect(() => send(source, owner.identity, 'context.request', { requestId: 'r2', scope: 'document' })).toThrow('permission denied');
    expect(() => send(owner, source.identity, 'context.response', { requestId: 'r1', snapshot: { text: 'private' } }, 'read')).toThrow('permission denied');
    expect(send(owner, source.identity, 'context.response', { requestId: 'r1', error: { code: 'ACCESS_DENIED' } }, 'read').result.state).toBe('responded');
    expect(() => send(source, owner.identity, 'task.submit', { taskId: 't1' })).toThrow('permission denied');
    const another = state.register({ workspaceId: 'w', documentId: 'other-peer' }, 'local');
    expect(send(another, owner.identity, 'task.submit', { taskId: 't2' }).result.state).toBe('submitted');
});

test('revoked context is redacted from snapshot recovery and event replay even after presence expires', () => {
    let now = 1000;
    const state = new CoordinationState({ clock: () => now });
    const owner = state.register({ workspaceId: 'w', documentId: 'owner' }, 'local');
    const peer = state.register({ workspaceId: 'w', documentId: 'peer' }, 'local');
    let count = 0;
    const send = (session, target, type, payload, correlationId = 'read') => state.append({ version: 2, workspaceId: 'w', source: session.identity,
        target, type, payload, createdAt: now, ttlMs: 900000, correlationId, idempotencyKey: `k-${++count}` }, session.credential, 'local');
    send(peer, owner.identity, 'context.request', { requestId: 'r', scope: 'document' });
    send(owner, peer.identity, 'context.response', { requestId: 'r', snapshot: { text: 'private reference' } });
    expect(JSON.stringify(state.snapshot(peer.credential, 'local'))).toContain('private reference');
    send(owner, owner.identity, 'document.announce', { access: { readContext: false } }, 'policy');
    now += 61000;
    expect(state.snapshot(peer.credential, 'local').documents).toEqual([]);
    expect(JSON.stringify(state.snapshot(peer.credential, 'local'))).not.toContain('private reference');
    expect(JSON.stringify(state.events(peer.credential, 'local', 0, state.epoch))).not.toContain('private reference');
    expect(() => send(peer, owner.identity, 'context.request', { requestId: 'another', scope: 'document' })).toThrow('permission denied');
});
