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
