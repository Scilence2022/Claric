import { createDistributedTaskRuntime } from '../src/taskpane/distributed-task-runtime.js';

describe('distributed task runtime', () => {
    function client() {
        return {
            identity: { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'instance-a' },
            submitTask: jest.fn(() => Promise.resolve()),
            cancelTask: jest.fn(() => Promise.resolve()),
        };
    }

    test('submits target-bound review-required tasks without runtime objects', async () => {
        const coordination = client();
        const runtime = createDistributedTaskRuntime({ client: coordination });
        await runtime.submitGraph({ graphId: 'graph-1', tasks: [{ taskId: 'task-1', type: 'edit', instruction: 'revise' }] }, () => ({
            workspaceId: 'workspace-a', documentId: 'document-b', instanceId: 'instance-b',
        }));
        expect(coordination.submitTask).toHaveBeenCalledWith(
            { workspaceId: 'workspace-a', documentId: 'document-b', instanceId: 'instance-b' },
            expect.objectContaining({ taskId: 'task-1', reviewRequired: true }),
            expect.objectContaining({ correlationId: expect.any(String), idempotencyKey: expect.any(String) }),
        );
        expect(JSON.stringify(coordination.submitTask.mock.calls[0])).not.toContain('Range');
    });

    test('accepts only workspace-scoped target events and deduplicates them', () => {
        const coordination = client();
        const statuses = [];
        const runtime = createDistributedTaskRuntime({ client: coordination, onStatus: (status) => statuses.push(status) });
        return runtime.submitGraph({ graphId: 'graph-2', tasks: [{ taskId: 'task-2', type: 'edit', instruction: 'revise' }] }, () => ({
            workspaceId: 'workspace-a', documentId: 'document-b', instanceId: 'instance-b',
        })).then(() => {
            const submitted = coordination.submitTask.mock.calls[0];
            const event = { eventId: 'event-1', workspaceId: 'workspace-a', type: 'task.progress', source: { workspaceId: 'workspace-a', documentId: 'document-b', instanceId: 'instance-b' }, target: coordination.identity, correlationId: submitted[2].correlationId, ttlMs: 30000, createdAt: Date.now(), payload: { taskId: 'task-2', graphId: 'graph-2', attemptId: runtime.getSnapshot('graph-2').tasks[0].attemptId, progress: 0.5 } };
            expect(runtime.handleEvent(event)).toBe(true);
            expect(runtime.handleEvent(event)).toBe(true);
            expect(statuses.filter((status) => status.type === 'task.progress')).toHaveLength(1);
            expect(runtime.handleEvent({ ...event, workspaceId: 'workspace-other' })).toBe(false);
        });
    });

    test('validates DAG before sends and only releases dependents after target success', async () => {
        const coordination = client();
        const target = { workspaceId: 'workspace-a', documentId: 'document-b', instanceId: 'instance-b' };
        const runtime = createDistributedTaskRuntime({ client: coordination, resolveTarget: () => target });
        await expect(runtime.submitGraph({ tasks: [{ taskId: 'a', dependsOn: ['missing'] }] })).rejects.toThrow('missing');
        await expect(runtime.submitGraph({ tasks: [{ taskId: 'a', dependsOn: ['b'] }, { taskId: 'b', dependsOn: ['a'] }] })).rejects.toThrow('cycle');
        expect(coordination.submitTask).not.toHaveBeenCalled();
        await runtime.submitGraph({ graphId: 'g', tasks: [{ taskId: 'a' }, { taskId: 'b', dependsOn: ['a'] }] });
        expect(coordination.submitTask).toHaveBeenCalledTimes(1);
        const [frozen, payload, metadata] = coordination.submitTask.mock.calls[0];
        expect(Object.isFrozen(frozen)).toBe(true);
        expect(payload).toMatchObject({ graphId: 'g', taskId: 'a', attemptId: expect.any(String) });
        expect(metadata).not.toHaveProperty('graphId');
        const event = { version: 2, workspaceId: 'workspace-a', source: { ...target }, target: coordination.identity, correlationId: metadata.correlationId, createdAt: Date.now(), ttlMs: 60000, payload: { taskId: 'a' } };
        target.documentId = 'mutated';
        expect(runtime.handleEvent({ ...event, type: 'task.succeeded', source: coordination.identity })).toBe(false);
        expect(runtime.handleEvent({ ...event, type: 'task.claim', idempotencyKey: 'claim' })).toBe(true);
        expect(coordination.submitTask).toHaveBeenCalledTimes(1);
        expect(runtime.handleEvent({ ...event, type: 'task.succeeded', idempotencyKey: 'success' })).toBe(true);
        await new Promise(setImmediate);
        expect(coordination.submitTask).toHaveBeenCalledTimes(2);
        expect(coordination.submitTask.mock.calls[1][0].documentId).toBe('document-b');
        await runtime.cancelGraph('g');
        expect(coordination.cancelTask).toHaveBeenCalledWith(expect.anything(), { taskId: 'b', reason: 'Cancelled by user' }, expect.objectContaining({ correlationId: expect.any(String) }));
        await expect(runtime.submitGraph({ graphId: 'g', tasks: [] })).rejects.toThrow('already');
        runtime.dispose();
    });
});

describe('review-aware distributed dependencies', () => {
    const local = { workspaceId: 'w', documentId: 'a', instanceId: 'a-1' };
    const target = { workspaceId: 'w', documentId: 'b', instanceId: 'b-1' };
    async function setup() {
        let now = 1000;
        const client = { identity: local, submitTask: jest.fn(async () => {}), cancelTask: jest.fn(async () => {}) };
        const runtime = createDistributedTaskRuntime({ client, clock: () => now });
        await runtime.submitGraph({ graphId: 'g', tasks: [{ taskId: 'a', target, ttlMs: 5000 }, { taskId: 'b', target, dependsOn: ['a'] }] });
        const metadata = client.submitTask.mock.calls[0][2];
        const send = (type, payload) => runtime.handleEvent({ type, payload, source: target, target: local, workspaceId: 'w',
            correlationId: metadata.correlationId, createdAt: now, ttlMs: 5000, idempotencyKey: type });
        send('task.claim', { taskId: 'a' });
        send('proposal.announced', { taskId: 'a', proposalId: 'p', revision: 0 });
        return { runtime, client, send, advance: () => { now += 5000; } };
    }
    test.each(['task.rejected', 'task.succeeded'])('rejection cannot release dependent writes, including legacy %s', async (terminal) => {
        const { runtime, client, send } = await setup();
        expect(runtime.getSnapshot('g').tasks[0].state).toBe('awaiting_review');
        send('proposal.decision', { proposalId: 'p', revision: 0, decision: 'rejected' });
        send(terminal, { taskId: 'a', proposalId: 'p' });
        await new Promise(setImmediate);
        expect(runtime.getSnapshot('g').tasks.map((task) => task.state)).toEqual(['rejected', 'blocked']);
        expect(client.submitTask).toHaveBeenCalledTimes(1);
    });
    test('expiry before review differs from uncertain writes and never retries', async () => {
        const waiting = await setup();
        waiting.advance(); waiting.runtime.expireTasks();
        expect(waiting.runtime.getSnapshot('g').tasks.map((task) => task.state)).toEqual(['expired', 'blocked']);
        const applying = await setup();
        applying.send('proposal.decision', { proposalId: 'p', revision: 0, decision: 'accepted' });
        expect(applying.runtime.getSnapshot('g').tasks[0].state).toBe('applying');
        applying.advance(); applying.runtime.expireTasks();
        expect(applying.runtime.getSnapshot('g').tasks.map((task) => task.state)).toEqual(['unknown', 'blocked']);
        expect(applying.client.submitTask).toHaveBeenCalledTimes(1);
    });
});
