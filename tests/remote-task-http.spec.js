/** @jest-environment node */
const http = require('http');
const { createCoordinationServer } = require('../scripts/coordination-server.cjs');
const { CoordinationState } = require('../src/lib/coordination/state.cjs');
const { createCoordinationClient } = require('../src/taskpane/coordination-client.js');
const { createRemoteTaskRunner } = require('../src/taskpane/remote-task-runner.js');

function httpFetch(url, options = {}) {
    return new Promise((resolve, reject) => {
        const req = http.request(url, { method: options.method || 'GET', headers: options.headers }, (res) => {
            const chunks = [];
            res.on('data', (chunk) => chunks.push(chunk));
            res.on('end', () => resolve({ ok: res.statusCode >= 200 && res.statusCode < 300, status: res.statusCode, json: async () => JSON.parse(Buffer.concat(chunks).toString()) }));
        });
        req.on('error', reject);
        req.end(options.body);
    });
}

describe('remote tasks through real HTTP v2', () => {
    let server, clients, source, target, now, captured, faults, runner;
    beforeEach(async () => {
        now = 1700000000000;
        clients = [];
        captured = [];
        faults = {};
        server = createCoordinationServer({ token: '', state: new CoordinationState({ clock: () => now }) });
        await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
        async function register(documentId) {
            const client = createCoordinationClient({
                identity: { workspaceId: 'w', documentId, instanceId: 'untrusted-local' },
                baseUrl: `http://127.0.0.1:${server.address().port}/coordination`, clock: () => now,
                setIntervalImpl: null,
                fetchImpl: async (url, options = {}) => {
                    const envelope = options.body ? JSON.parse(options.body) : null;
                    if (envelope?.type) captured.push(envelope);
                    if (faults.before) await faults.before(envelope, url);
                    const response = await httpFetch(url, options);
                    if (faults.after) await faults.after(envelope, response);
                    return response;
                },
            });
            clients.push(client);
            await client.start();
            expect(client.identity.instanceId).not.toBe('untrusted-local');
            return client;
        }
        source = await register('a');
        target = await register('b');
    });
    afterEach(async () => {
        faults = {};
        runner?.dispose();
        await Promise.all(clients.map((client) => client.stop()));
        await new Promise((resolve) => server.close(resolve));
    });
    async function prepare(overrides = {}) {
        let revision = 'r1';
        const localResult = { applied: true, count: 1 };
        const agent = {
            prepareTask: jest.fn(async () => ({ kind: 'edit', scope: 'selection', title: 'Review', summary: 'Target-local edit', baseRevision: 'r1', items: [{ id: 'one', before: 'private original', after: 'private replacement' }] })),
            getDocumentRevision: jest.fn(async () => revision),
            applyProposal: jest.fn(async () => { revision = 'r2'; return localResult; }),
            ...overrides.agent,
        };
        let review;
        runner = createRemoteTaskRunner({ client: target, documentAgent: agent, clock: () => now, ...overrides.runner, onProposal: (record, runtime) => { review = { record, runtime }; } });
        const sent = await source.submitTask(target.identity, { taskId: 't1', taskType: 'edit', instruction: 'Revise', graphId: 'g1', attemptId: 'a1' }, { correlationId: 'c1', idempotencyKey: 'submit1', ttlMs: 60000, createdAt: now });
        const snapshot = await target.refresh();
        const event = snapshot.events.find((entry) => entry.type === 'task.submit');
        expect(event).toEqual(sent.event);
        await Promise.all([runner.handleEvent(event), runner.handleEvent(event)]);
        expect(review).toBeDefined();
        return { agent, review, event, localResult };
    }
    test.each(['apply', 'reject'])('registers, claims, prepares and %s reaches source terminal state', async (decision) => {
        const { agent, review, event } = await prepare();
        expect(agent.prepareTask).toHaveBeenCalledWith(expect.objectContaining({ taskId: 't1', instruction: 'Revise' }), { signal: expect.any(AbortSignal) });
        expect(review.record.source).toEqual(source.identity);
        expect(review.record.target).toEqual(target.identity);
        expect(review.record.items[0].after).toBe('private replacement');
        const announcement = captured.find((entry) => entry.type === 'proposal.announced');
        expect(JSON.stringify(announcement)).not.toMatch(/private original|private replacement|"items"|"proposal":/);
        expect(announcement.correlationId).toBe(event.correlationId);
        expect(announcement.createdAt + announcement.ttlMs).toBe(event.createdAt + event.ttlMs);
        const [result] = await Promise.all([review.runtime[decision](), review.runtime[decision]()]);
        expect(result).toMatchObject({ ok: true, record: { state: decision === 'apply' ? 'applied' : 'rejected' } });
        expect(agent.applyProposal).toHaveBeenCalledTimes(decision === 'apply' ? 1 : 0);
        await runner.handleEvent(event);
        expect(agent.prepareTask).toHaveBeenCalledTimes(1);
        await review.runtime[decision]();
        expect(agent.applyProposal).toHaveBeenCalledTimes(decision === 'apply' ? 1 : 0);
        const received = await source.refresh();
        expect(received.events.some((entry) => entry.type === 'task.succeeded')).toBe(true);
        const final = await source.transport.getSnapshot();
        expect(final.tasks[0].state).toBe('succeeded');
        expect(final.proposals[0]).toMatchObject({ state: decision === 'apply' ? 'applied' : 'rejected', source: source.identity, target: target.identity });
        if (decision === 'apply') {
            expect(final.proposals[0].documentRevision).toBe('r2');
            expect(final.proposals[0].fence).toBeGreaterThan(0);
            expect(captured.findIndex((entry) => entry.type === 'proposal.decision')).toBeLessThan(captured.findIndex((entry) => entry.type === 'proposal.applied'));
            expect(captured.find((entry) => entry.type === 'lease.acquire').payload.resourceId).toBe('write');
        }
    });
    test('server rejects wrong actors, missing claim, nested proposal and zero fences', async () => {
        const sent = await source.submitTask(target.identity, { taskId: 't1', instruction: 'Revise' }, { correlationId: 'c1', createdAt: now, ttlMs: 60000 });
        const announce = { type: 'proposal.announced', target: source.identity, correlationId: 'c1', createdAt: now, ttlMs: 60000, payload: { proposalId: 'p1', taskId: 't1', revision: 0, baseRevision: 'r1', artifactIds: [] } };
        await expect(target.publishEnvelope(announce)).rejects.toMatchObject({ status: 409 });
        await target.publishEnvelope({ type: 'task.claim', target: source.identity, correlationId: sent.event.correlationId, payload: { taskId: 't1' } });
        await expect(source.publishEnvelope({ ...announce, target: target.identity })).rejects.toMatchObject({ status: 403 });
        await expect(target.publishEnvelope({ ...announce, payload: { ...announce.payload, proposal: { items: [] } } })).rejects.toMatchObject({ status: 400 });
        await target.publishEnvelope(announce);
        await expect(target.publishEnvelope({ type: 'proposal.decision', target: source.identity, correlationId: 'c1', payload: { proposalId: 'p1', revision: 0, decision: 'accepted' } })).rejects.toMatchObject({ status: 409 });
        await expect(target.publishEnvelope({ type: 'lease.release', target: target.identity, payload: { resourceId: 'write', fence: 0 } })).rejects.toMatchObject({ status: 400 });
    });
    test.each(['cancel', 'expired', 'offline', 'stale'])('%s refuses a new write', async (mode) => {
        const { review, agent } = await prepare();
        if (mode === 'cancel') await source.publishEnvelope({ type: 'task.cancel', target: target.identity, correlationId: 'c1', payload: { taskId: 't1' } });
        if (mode === 'expired') now += 60001;
        if (mode === 'offline') faults.before = async (envelope) => { if (envelope?.type === 'lease.acquire') throw new Error('Connection lost'); };
        if (mode === 'stale') agent.getDocumentRevision.mockResolvedValue('r2');
        expect((await review.runtime.apply()).ok).toBe(false);
        expect(agent.applyProposal).not.toHaveBeenCalled();
    });
    test('cancellation after accepted decision but before write refuses mutation', async () => {
        const { review, agent } = await prepare();
        faults.after = async (envelope) => {
            if (envelope?.type !== 'proposal.decision') return;
            faults.after = null;
            await source.publishEnvelope({ type: 'task.cancel', target: target.identity, correlationId: 'c1', payload: { taskId: 't1' } });
        };
        expect((await review.runtime.apply()).ok).toBe(false);
        expect(agent.applyProposal).not.toHaveBeenCalled();
    });
    test('lost acknowledgment after local write preserves local result and blocks retry', async () => {
        const { review, agent, localResult } = await prepare();
        faults.after = async (envelope) => { if (envelope?.type === 'proposal.applied') throw new Error('Lost response'); };
        const result = await review.runtime.apply();
        expect(result).toMatchObject({ ok: false, record: { state: 'unknown', items: [{ status: 'applied' }] }, localResult });
        expect(review.runtime.getLocalResult()).toEqual(localResult);
        await review.runtime.apply();
        expect(agent.applyProposal).toHaveBeenCalledTimes(1);
    });
    test('uses the actual document agent for preparation, review and local application over HTTP', async () => {
        const { createDocumentAgent } = require('../src/taskpane/document-agent.js');
        const { webcrypto } = require('crypto');
        let text = 'Original target passage';
        const model = jest.fn(async () => ({ selectionText: text, amendedText: 'Revised target passage' }));
        const write = jest.fn(async (deps, proposal) => { text = proposal.amendedText; });
        const agent = createDocumentAgent({
            identity: target.identity, clock: () => now, cryptoImpl: webcrypto,
            actions: { readSelectionContent: async () => ({ text }), prepareSelectionAmendment: model, applySelectionAmendment: write },
        });
        let review;
        runner = createRemoteTaskRunner({ client: target, documentAgent: agent, clock: () => now,
            onProposal: (record, runtime) => { review = { record, runtime }; } });
        await source.submitTask(target.identity, { taskId: 'agent-task', taskType: 'edit', instruction: 'Polish this text' },
            { correlationId: 'agent-correlation', ttlMs: 60000, createdAt: now });
        const incoming = await target.refresh();
        await runner.handleEvent(incoming.events.find((event) => event.type === 'task.submit'));
        expect(model).toHaveBeenCalledTimes(1);
        expect(write).not.toHaveBeenCalled();
        expect(review.record.items[0].after).toBe('Revised target passage');
        expect((await review.runtime.apply(['text-1'])).ok).toBe(true);
        expect(text).toBe('Revised target passage');
        expect((await source.transport.getSnapshot()).tasks[0].state).toBe('succeeded');
    });
    test('local write exception is unknown and cannot be retried', async () => {
        const { review, agent } = await prepare({ agent: { applyProposal: jest.fn(async () => { throw new Error('Word sync lost'); }) } });
        expect(await review.runtime.apply()).toMatchObject({ ok: false, record: { state: 'unknown' } });
        await review.runtime.apply();
        expect(agent.applyProposal).toHaveBeenCalledTimes(1);
        expect((await source.transport.getSnapshot()).tasks[0].state).toBe('failed');
    });
});
