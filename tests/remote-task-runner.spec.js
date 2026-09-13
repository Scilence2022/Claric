import { createRemoteTaskRunner } from '../src/taskpane/remote-task-runner.js';
const { CoordinationState } = require('../src/lib/coordination/state.cjs');

function setup(overrides = {}) {
    const state = new CoordinationState({ clock: () => 1000 });
    const binding = 'test';
    const sender = state.register({ workspaceId: 'w', documentId: 'a' }, binding);
    const receiver = state.register({ workspaceId: 'w', documentId: 'b' }, binding);
    const identity = receiver.identity;
    const published = [];
    const client = {
        identity,
        publishEnvelope: jest.fn(async (input) => {
            published.push(input);
            return { event: state.append({ version: 2, workspaceId: 'w', source: identity, ...input }, receiver.credential, binding) };
        }),
        getSnapshot: async () => state.snapshot(receiver.credential, binding),
    };
    const documentAgent = { getDocumentRevision: jest.fn(async () => 'rev-1'), applyProposal: jest.fn(async () => ({ ok: true })) };
    const onProposal = jest.fn();
    const runner = createRemoteTaskRunner({
        client, identity, documentAgent,
        prepareTask: jest.fn(async () => ({ kind: 'replace', scope: 'selection', title: 'T', summary: 'S', items: [{ id: 'one' }], baseRevision: 'rev-1' })),
        onProposal, clock: () => 1000, ...overrides,
    });
    const event = state.append({ version: 2, workspaceId: 'w', source: sender.identity, target: identity, type: 'task.submit', createdAt: 1000, ttlMs: 60000, correlationId: 'c1', idempotencyKey: 'i1', payload: { taskId: 't1', graphId: 'g1', attemptId: 'a1' } }, sender.credential, binding);
    return { runner, client, documentAgent, onProposal, published, event, identity };
}

describe('remote task runner', () => {
    test('filters exact target and returns an awaitable event handler', async () => {
        const { runner, event, identity } = setup();
        for (const key of ['workspaceId', 'documentId', 'instanceId']) {
            expect(await runner.handleEvent({ ...event, target: { ...identity, [key]: 'other' } })).toBe(false);
        }
    });
    test('deduplicates concurrent and later submissions with default createProposal', async () => {
        const { runner, onProposal, published, event, identity } = setup();
        await Promise.all([runner.handleEvent(event), runner.handleEvent(event)]);
        await runner.handleEvent(event);
        expect(published.map((item) => item.type)).toEqual(['task.claim', 'proposal.announced']);
        expect(onProposal).toHaveBeenCalledTimes(1);
        expect(onProposal.mock.calls[0][0]).toMatchObject({ reviewRequired: true, target: identity, items: [{ id: 'one' }] });
    });
    test('rejects callback, range, and full-text prepared values', async () => {
        for (const bad of [{ callback: () => {} }, { range: { load() {} } }, { fullText: 'secret' }]) {
            const { runner, published, event } = setup({ prepareTask: async () => bad });
            await runner.handleEvent(event);
            expect(published.some((item) => item.type === 'task.failed')).toBe(true);
        }
    });
    test('requires a real local apply method', () => {
        expect(() => setup({ documentAgent: { getDocumentRevision: async () => 'r1' } })).toThrow('applyProposal');
    });
    test('server decision and lease are accepted before the document agent writes', async () => {
        const { runner, onProposal, published, event, documentAgent } = setup();
        documentAgent.applyProposal.mockImplementation(async () => {
            expect(published.map((item) => item.type)).toEqual(['task.claim', 'proposal.announced', 'lease.acquire', 'proposal.decision', 'lease.renew']);
            return { ok: true };
        });
        await runner.handleEvent(event);
        expect((await onProposal.mock.calls[0][1].apply()).ok).toBe(true);
        expect(published.map((item) => item.type)).toEqual(['task.claim', 'proposal.announced', 'lease.acquire', 'proposal.decision', 'lease.renew', 'proposal.applied', 'task.succeeded', 'lease.release']);
    });
    test('does not expose a proposal when server announcement is rejected', async () => {
        const { runner, client, onProposal, event } = setup();
        const publish = client.publishEnvelope.getMockImplementation();
        client.publishEnvelope.mockImplementation((input) => {
            if (input.type === 'proposal.announced') throw new Error('Rejected');
            return publish(input);
        });
        await runner.handleEvent(event);
        expect(onProposal).not.toHaveBeenCalled();
    });
});
