import { createContextRequestManager } from '../src/taskpane/context-requests.js';

function makeClient() {
    return {
        identity: { workspaceId: 'workspace-a', documentId: 'document-b', instanceId: 'instance-b' },
        requestContext: jest.fn(() => Promise.resolve()),
        respondContext: jest.fn(() => Promise.resolve()),
    };
}

function event(type, source, target, payload, correlationId = 'correlation-1') {
    return { type, source, target, workspaceId: source.workspaceId, payload, correlationId, ttlMs: 30000, createdAt: Date.now() };
}

describe('context request manager', () => {
    test('responds to a targeted request and ignores another document', async () => {
        const client = makeClient();
        const agent = { readContext: jest.fn(async () => ({ text: 'bounded data', snapshotId: 'snapshot-1' })) };
        const manager = createContextRequestManager({ client, agent });
        const target = client.identity;
        const source = { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'instance-a' };
        expect(await manager.handleEvent(event('context.request', source, target, { requestId: 'request-1', scope: 'selection' }))).toBe(true);
        expect(client.respondContext).toHaveBeenCalledTimes(1);
        expect(agent.readContext).toHaveBeenCalledWith({ requestId: 'request-1', scope: 'selection' }, expect.objectContaining({ signal: expect.any(AbortSignal) }));
        expect(await manager.handleEvent(event('context.request', source, { ...target, documentId: 'document-c' }, { requestId: 'request-2', scope: 'selection' }, 'correlation-2'))).toBe(false);
    });

    test('resolves a request once and ignores duplicate response', async () => {
        const client = makeClient();
        const manager = createContextRequestManager({ client, agent: { readContext: jest.fn() }, timeoutMs: 1000 });
        const target = { workspaceId: 'workspace-a', documentId: 'document-a', instanceId: 'instance-a' };
        const promise = manager.requestContext(target, { scope: 'selection' }, { correlationId: 'correlation-9' });
        const request = client.requestContext.mock.calls[0][0];
        await manager.handleEvent(event('context.response', target, client.identity, { requestId: client.requestContext.mock.calls[0][1].requestId, snapshot: { text: 'result', sourceDocumentId: target.documentId, sourceInstanceId: target.instanceId } }, 'correlation-9'));
        expect(await promise).toMatchObject({ snapshot: { text: 'result' } });
        expect(await manager.handleEvent(event('context.response', target, client.identity, { requestId: client.requestContext.mock.calls[0][1].requestId, text: 'late' }, 'correlation-9'))).toBe(true);
        expect(request.documentId).toBe('document-a');
        manager.dispose();
    });

    test('real state accepts cross-document request and exact duplicate response envelopes', async () => {
        const { CoordinationState } = require('../src/lib/coordination/state.cjs');
        const { createCoordinationClient } = require('../src/taskpane/coordination-client.js');
        let now = 100000;
        const state = new CoordinationState({ clock: () => now });
        const wire = [];
        const fetchImpl = async (url, options = {}) => {
            const credential = options.headers.Authorization?.slice(7);
            const body = options.body ? JSON.parse(options.body) : undefined;
            try {
                let result;
                if (url.endsWith('/register')) result = state.register(body, 'binding');
                else if (url.endsWith('/envelopes')) { wire.push(body); result = { event: state.append(body, credential, 'binding') }; }
                else if (url.endsWith('/snapshot')) result = state.snapshot(credential, 'binding');
                else { const query = new URL(url, 'http://localhost').searchParams; result = state.events(credential, 'binding', Number(query.get('after')), query.get('epoch')); }
                return { ok: true, json: async () => result };
            } catch (error) { return { ok: false, status: error.status || 400, json: async () => ({ error: error.message }) }; }
        };
        const a = createCoordinationClient({ fetchImpl, clock: () => now, setIntervalImpl: null, identity: { workspaceId: 'w', documentId: 'a', instanceId: 'a' } });
        const b = createCoordinationClient({ fetchImpl, clock: () => now, setIntervalImpl: null, identity: { workspaceId: 'w', documentId: 'b', instanceId: 'b' } });
        await a.start(); await b.start();
        const readContext = jest.fn(async () => ({ text: 'from B', sourceDocumentId: 'b', sourceInstanceId: b.identity.instanceId }));
        const ma = createContextRequestManager({ client: a, agent: { readContext }, clock: () => now });
        const mb = createContextRequestManager({ client: b, agent: { readContext }, clock: () => now });
        try {
            const answer = ma.requestContext(b.identity, { requestId: 'r', scope: 'selection', maxChars: 10, maxTokens: 3 }, { correlationId: 'c' });
            await new Promise(setImmediate);
            const request = wire.find((item) => item.type === 'context.request');
            expect(request.payload).toEqual({ requestId: 'r', scope: 'selection' });
            await mb.handleEvent(request);
            now += 100;
            await mb.handleEvent(request);
            const responses = wire.filter((item) => item.type === 'context.response');
            expect(responses).toHaveLength(2);
            expect(responses[1]).toEqual(responses[0]);
            expect(readContext).toHaveBeenCalledTimes(1);
            expect(await ma.handleEvent({ ...responses[0], source: { ...b.identity, documentId: 'forged' } })).toBe(false);
            await ma.handleEvent(responses[0]);
            expect(await answer).toMatchObject({ requestId: 'r', snapshot: { text: 'from B' } });
        } finally { ma.dispose(); mb.dispose(); await a.stop(); await b.stop(); }
    });

    test('abort notifies peer and cancels the in-flight local read', async () => {
        const client = makeClient();
        let signal;
        let finishRead;
        const manager = createContextRequestManager({ client, agent: { readContext: (_payload, options) => { signal = options.signal; return new Promise((resolve) => { finishRead = resolve; }); } } });
        const source = { workspaceId: 'workspace-a', documentId: 'a', instanceId: 'a' };
        const request = event('context.request', source, client.identity, { requestId: 'r', scope: 'selection' });
        const first = manager.handleEvent(request);
        await Promise.resolve();
        const duplicate = manager.handleEvent(request);
        await manager.handleEvent({ ...request, type: 'context.cancel' });
        expect(signal.aborted).toBe(true);
        finishRead({ text: 'late' });
        await Promise.all([first, duplicate]);
        expect(client.respondContext).not.toHaveBeenCalled();
        manager.dispose();
    });
});
