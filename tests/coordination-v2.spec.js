const fs = require('fs');
const path = require('path');
const { CoordinationState } = require('../src/lib/coordination/state.cjs');
const { EVENT_TYPES, PROTOCOL_V2 } = require('../src/lib/coordination/protocol.cjs');

const now = 1700000000000;
const binding = 'loopback|http://localhost:3010';
function envelope(source, target, type, payload, id) {
    return {
        version: PROTOCOL_V2,
        workspaceId: source.workspaceId,
        source,
        target,
        type,
        correlationId: `correlation-${id}`,
        idempotencyKey: `idempotency-${id}`,
        ttlMs: 30000,
        createdAt: now,
        payload,
    };
}

describe('coordination v2 state', () => {
    let state;
    let a;
    let b;

    beforeEach(() => {
        state = new CoordinationState({ clock: () => now });
        a = state.register({ workspaceId: 'workspace-a', documentId: 'document-a' }, binding);
        b = state.register({ workspaceId: 'workspace-a', documentId: 'document-b' }, binding);
    });

    test('issues server identities and authenticates the issued credential', () => {
        expect(a.identity.instanceId).not.toBe(b.identity.instanceId);
        expect(() => state.authenticate(a.credential, 'different-binding')).toThrow('Invalid or expired');
        expect(state.authenticate(a.credential, binding)).toEqual(a.identity);
    });

    test('routes a document announce and context request/response to the bound instances', () => {
        const announce = envelope(a.identity, a.identity, EVENT_TYPES.DOCUMENT_ANNOUNCE, {
            title: 'Document A', capabilities: ['readContext'],
        }, 'announce');
        expect(state.append(announce, a.credential, binding).result.documentId).toBe('document-a');

        const request = envelope(a.identity, b.identity, EVENT_TYPES.CONTEXT_REQUEST, {
            requestId: 'request-1', scope: 'document',
        }, 'request');
        expect(state.append(request, a.credential, binding).result.state).toBe('pending');

        const response = envelope(b.identity, a.identity, EVENT_TYPES.CONTEXT_RESPONSE, {
            requestId: 'request-1', artifactIds: ['artifact-1'], summary: 'bounded result',
        }, 'request');
        expect(state.append(response, b.credential, binding).result.state).toBe('responded');
        expect(state.snapshot(a.credential, binding).requests[0].state).toBe('responded');
    });

    test('makes task claim and lease ownership explicit', () => {
        const task = envelope(a.identity, b.identity, EVENT_TYPES.TASK_SUBMIT, {
            taskId: 'task-1', taskType: 'edit', instruction: 'revise target', reviewRequired: true,
        }, 'task');
        state.append(task, a.credential, binding);
        const claim = { ...envelope(b.identity, a.identity, EVENT_TYPES.TASK_CLAIM, { taskId: 'task-1' }, 'task'), correlationId: task.correlationId };
        expect(state.append(claim, b.credential, binding).result.state).toBe('claimed');
        const lease = envelope(b.identity, b.identity, EVENT_TYPES.LEASE_ACQUIRE, { resourceId: 'write', durationMs: 5000 }, 'lease');
        const leaseEvent = state.append(lease, b.credential, binding);
        expect(leaseEvent.result.fence).toBe(1);
        expect(() => state.append({ ...lease, idempotencyKey: 'lease-stale' }, a.credential, binding)).toThrow('Session does not own');
    });

    test('deduplicates an idempotency key and exposes a replay cursor', () => {
        const event = envelope(a.identity, b.identity, EVENT_TYPES.CONTEXT_REQUEST, { requestId: 'request-2', scope: 'document' }, 'same');
        const first = state.append(event, a.credential, binding);
        const second = state.append(event, a.credential, binding);
        expect(second).toEqual(first);
        const snapshot = state.snapshot(a.credential, binding);
        expect(state.events(a.credential, binding, 0, snapshot.epoch).events).toHaveLength(1);
    });
});

describe('coordination persistence', () => {
    const checkpointDir = path.join(process.cwd(), 'tmp-coordination-test-data');
    const checkpointPath = path.join(checkpointDir, 'coordination-checkpoint.json');
    const persistedBinding = 'loopback|http://localhost:3010';

    beforeEach(() => {
        fs.rmSync(checkpointDir, { recursive: true, force: true });
        fs.mkdirSync(checkpointDir, { mode: 0o700 });
        fs.chmodSync(checkpointDir, 0o700);
    });
    afterEach(() => fs.rmSync(checkpointDir, { recursive: true, force: true }));

    test('recovers non-terminal tasks as interrupted', () => {
        const first = new CoordinationState({ clock: () => now, persistencePath: checkpointPath });
        const a = first.register({ workspaceId: 'workspace-p', documentId: 'document-a' }, persistedBinding);
        const b = first.register({ workspaceId: 'workspace-p', documentId: 'document-b' }, persistedBinding);
        first.append(envelope(a.identity, b.identity, EVENT_TYPES.TASK_SUBMIT, { taskId: 'task-p', taskType: 'edit', instruction: 'revise' }, 'task-p'), a.credential, persistedBinding);
        const recovered = new CoordinationState({ clock: () => now, persistencePath: checkpointPath });
        expect(recovered.recovery.tasks[0]).toMatchObject({ taskId: 'task-p', state: 'interrupted', previousState: 'submitted' });
    });

    test('blocks writes after an external checkpoint modification', () => {
        const first = new CoordinationState({ clock: () => now, persistencePath: checkpointPath });
        const a = first.register({ workspaceId: 'workspace-p', documentId: 'document-a' }, persistedBinding);
        const b = first.register({ workspaceId: 'workspace-p', documentId: 'document-b' }, persistedBinding);
        first.append(envelope(a.identity, b.identity, EVENT_TYPES.CONTEXT_REQUEST, { requestId: 'request-p', scope: 'document' }, 'request-p'), a.credential, persistedBinding);
        fs.appendFileSync(checkpointPath, ' ');
        expect(() => first.append(envelope(a.identity, b.identity, EVENT_TYPES.CONTEXT_REQUEST, { requestId: 'request-p2', scope: 'document' }, 'request-p2'), a.credential, persistedBinding)).toThrow('Checkpoint failed');
        expect(() => first.append(envelope(a.identity, b.identity, EVENT_TYPES.CONTEXT_REQUEST, { requestId: 'request-p3', scope: 'document' }, 'request-p3'), a.credential, persistedBinding)).toThrow('restart required');
    });
});
