const { connectMcpServer } = require('../src/lib/mcp-client.js');
const { runToolLoop } = require('../src/lib/tool-loop.js');
const { createMcpToolExecutor, createResourceClient, importServerPrompts } = require('../src/lib/mcp-tools.js');

function deferred() {
    let resolve;
    let reject;
    const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
    return { promise, resolve, reject };
}

function response(result = {}) {
    return { ok: true, headers: { get: () => null }, text: async () => JSON.stringify({ result }) };
}

async function pendingBody(options = {}) {
    const body = deferred();
    const started = deferred();
    const fetchFn = jest.fn().mockImplementation(async () => ({
        ...response(),
        text: () => { started.resolve(); return body.promise; },
    }));
    const pending = connectMcpServer({ url: '/mcp', fetchFn, timeoutMs: 100, ...options });
    const outcome = pending.then((value) => ({ value }), (error) => ({ error }));
    await started.promise;
    return { body, fetchFn, outcome };
}

describe('MCP request deadlines and cancellation', () => {
    beforeEach(() => jest.useFakeTimers());
    afterEach(() => { jest.useRealTimers(); jest.restoreAllMocks(); });

    test.each(['application/json', 'text/event-stream'])('deadline includes deferred %s body', async (contentType) => {
        const body = deferred();
        const started = deferred();
        const fetchFn = jest.fn(async () => ({
            ...response(),
            headers: { get: () => contentType },
            text: () => { started.resolve(); return body.promise; },
        }));
        const outcome = connectMcpServer({ url: '/mcp', fetchFn, timeoutMs: 100 }).catch((error) => error);
        await started.promise;
        await jest.advanceTimersByTimeAsync(99);
        expect(fetchFn.mock.calls[0][1].signal.aborted).toBe(false);
        await jest.advanceTimersByTimeAsync(1);
        expect(await outcome).toMatchObject({ name: 'AbortError', message: 'MCP request timed out.' });
        expect(fetchFn.mock.calls[0][1].signal.aborted).toBe(true);
        expect(jest.getTimerCount()).toBe(0);
        body.reject(new Error('late body failure'));
        await Promise.resolve();
    });

    test.each(['deadline', 'signal'])('real Response stream body is bounded by %s', async (mode) => {
        const controller = new AbortController();
        const started = deferred();
        let streamController;
        const stream = new ReadableStream({ start(value) { streamController = value; } });
        const wireResponse = new Response(stream, { headers: { 'content-type': 'application/json' } });
        const text = wireResponse.text.bind(wireResponse);
        jest.spyOn(wireResponse, 'text').mockImplementation(() => { started.resolve(); return text(); });
        const fetchFn = jest.fn(async () => wireResponse);
        const outcome = connectMcpServer({ url: '/mcp', fetchFn, timeoutMs: 100, signal: controller.signal })
            .catch((error) => error);
        await started.promise;
        streamController.enqueue(new TextEncoder().encode('{"result":'));
        if (mode === 'signal') controller.abort();
        else await jest.advanceTimersByTimeAsync(100);
        expect(await outcome).toMatchObject({ name: 'AbortError' });
        expect(fetchFn.mock.calls[0][1].signal.aborted).toBe(true);
        expect(jest.getTimerCount()).toBe(0);
        streamController.close();
    });

    test('external abort rejects an uncooperative body and removes its listener', async () => {
        const controller = new AbortController();
        const add = jest.spyOn(controller.signal, 'addEventListener');
        const remove = jest.spyOn(controller.signal, 'removeEventListener');
        const { outcome, body, fetchFn } = await pendingBody({ signal: controller.signal });
        controller.abort();
        expect((await outcome).error).toMatchObject({ name: 'AbortError' });
        expect(fetchFn.mock.calls[0][1].signal.aborted).toBe(true);
        expect(remove).toHaveBeenCalledWith('abort', add.mock.calls[0][1]);
        expect(jest.getTimerCount()).toBe(0);
        body.resolve('{}');
    });

    test('pre-aborted signal does not fetch or schedule a timer', async () => {
        const controller = new AbortController();
        controller.abort();
        const fetchFn = jest.fn();
        await expect(connectMcpServer({ url: '/mcp', fetchFn, signal: controller.signal }))
            .rejects.toMatchObject({ name: 'AbortError' });
        expect(fetchFn).not.toHaveBeenCalled();
        expect(jest.getTimerCount()).toBe(0);
    });

    test('external abort also bounds waiting for headers', async () => {
        const controller = new AbortController();
        const headers = deferred();
        const started = deferred();
        const fetchFn = jest.fn(() => { started.resolve(); return headers.promise; });
        const outcome = connectMcpServer({ url: '/mcp', fetchFn, signal: controller.signal }).catch((error) => error);
        await started.promise;
        controller.abort();
        expect(await outcome).toMatchObject({ name: 'AbortError' });
        expect(jest.getTimerCount()).toBe(0);
        headers.reject(new Error('late fetch failure'));
    });

    test.each(['Error', 'AbortError'])('body %s propagates and cleans the deadline and listener', async (name) => {
        const controller = new AbortController();
        const remove = jest.spyOn(controller.signal, 'removeEventListener');
        const { outcome, body } = await pendingBody({ signal: controller.signal });
        body.reject(Object.assign(new Error('body read failed'), { name }));
        expect((await outcome).error).toMatchObject({ name, message: 'body read failed' });
        expect(remove).toHaveBeenCalledTimes(1);
        expect(jest.getTimerCount()).toBe(0);
    });

    test.each(['fetch', 'http', 'rpc'])('%s failure cleans the deadline and external listener', async (kind) => {
        const controller = new AbortController();
        const remove = jest.spyOn(controller.signal, 'removeEventListener');
        const fetchFn = jest.fn(async () => {
            if (kind === 'fetch') throw new Error('transport failed');
            if (kind === 'http') return { ok: false, status: 500 };
            return { ...response(), text: async () => JSON.stringify({ error: { code: -1, message: 'rpc failed' } }) };
        });
        await expect(connectMcpServer({ url: '/mcp', fetchFn, signal: controller.signal })).rejects.toThrow();
        expect(remove).toHaveBeenCalledTimes(1);
        expect(jest.getTimerCount()).toBe(0);
    });

    test('successful requests release timers/listeners and per-call signals leave the client reusable', async () => {
        const defaultController = new AbortController();
        const add = jest.spyOn(defaultController.signal, 'addEventListener');
        const remove = jest.spyOn(defaultController.signal, 'removeEventListener');
        const fetchFn = jest.fn(async () => response());
        const client = await connectMcpServer({ url: '/mcp', fetchFn, signal: defaultController.signal });
        expect(remove).toHaveBeenCalledTimes(2);
        expect(add).toHaveBeenCalledTimes(2);
        const requestController = new AbortController();
        const body = deferred();
        const started = deferred();
        fetchFn.mockImplementationOnce(async () => ({
            ...response(), text: () => { started.resolve(); return body.promise; },
        }));
        const outcome = client.callTool('read', {}, { signal: requestController.signal }).catch((error) => error);
        await started.promise;
        requestController.abort();
        expect(await outcome).toMatchObject({ name: 'AbortError' });
        expect(defaultController.signal.aborted).toBe(false);
        await expect(client.listTools()).resolves.toEqual([]);
        expect(jest.getTimerCount()).toBe(0);
        body.resolve('{}');
    });

    test.each(['listPrompts', 'listResources'])('%s does not mistake AbortError for unsupported methods', async (method) => {
        const fetchFn = jest.fn(async () => response());
        const client = await connectMcpServer({ url: '/mcp', fetchFn });
        fetchFn.mockRejectedValueOnce(new DOMException('Method not found', 'AbortError'));
        await expect(client[method]()).rejects.toMatchObject({ name: 'AbortError' });
    });
});

const loopArgs = { systemPrompt: '', taskPrompt: '', tools: [{ name: 'read' }], maxSteps: 1 };

describe('tool-loop cancellation boundaries', () => {
    test.each(['read', 'finish', 'invalid'])('abort during deferred send rejects before handling %s', async (tool) => {
        const controller = new AbortController();
        const reply = deferred();
        const execute = jest.fn();
        const onStep = jest.fn();
        const outcome = runToolLoop({ ...loopArgs, send: () => reply.promise, execute, onStep, signal: controller.signal });
        controller.abort();
        reply.resolve(tool === 'invalid' ? 'not json' : JSON.stringify({ tool }));
        await expect(outcome).rejects.toMatchObject({ name: 'AbortError' });
        expect(execute).not.toHaveBeenCalled();
        expect(onStep).not.toHaveBeenCalled();
    });

    test.each(['resolve', 'reject'])('abort during deferred execute %s cannot become step-limit', async (settle) => {
        const controller = new AbortController();
        const toolResult = deferred();
        const started = deferred();
        const onStep = jest.fn();
        const send = jest.fn(async () => '{"tool":"read"}');
        const outcome = runToolLoop({ ...loopArgs, send, onStep, signal: controller.signal,
            execute: () => { started.resolve(); return toolResult.promise; },
        });
        await started.promise;
        controller.abort();
        if (settle === 'resolve') toolResult.resolve({ ok: true });
        else toolResult.reject(new Error('transport interrupted'));
        await expect(outcome).rejects.toMatchObject({ name: 'AbortError' });
        expect(onStep).toHaveBeenCalledTimes(1);
        expect(send).toHaveBeenCalledTimes(1);
    });

    test('tool AbortError propagates unchanged without an external signal', async () => {
        const error = new DOMException('tool aborted', 'AbortError');
        const result = deferred();
        const started = deferred();
        const send = jest.fn(async () => '{"tool":"read"}');
        const outcome = runToolLoop({ ...loopArgs, send,
            execute: () => { started.resolve(); return result.promise; },
        });
        await started.promise;
        result.reject(error);
        await expect(outcome).rejects.toBe(error);
        expect(send).toHaveBeenCalledTimes(1);
    });

    test.each(['reply', 'finish', 'observation'])('abort inside %s hook does not return success', async (phase) => {
        const controller = new AbortController();
        const execute = jest.fn(async () => ({ ok: true }));
        const outcome = runToolLoop({ ...loopArgs, execute, signal: controller.signal,
            send: async () => JSON.stringify({ tool: phase === 'finish' ? 'finish' : 'read' }),
            onStep: ({ call }) => {
                if ((phase === 'reply' && !call) || (phase !== 'reply' && call)) controller.abort();
            },
        });
        await expect(outcome).rejects.toMatchObject({ name: 'AbortError' });
        if (phase !== 'observation') expect(execute).not.toHaveBeenCalled();
    });
});

describe('MCP bridge cancellation', () => {
    test('deferred client AbortError crosses executor and loop unchanged', async () => {
        const result = deferred();
        const started = deferred();
        const error = new DOMException('cancelled', 'AbortError');
        const execute = createMcpToolExecutor(new Map([['read', {
            originalName: 'read', client: { callTool: () => { started.resolve(); return result.promise; } },
        }]]));
        const outcome = runToolLoop({ ...loopArgs, execute, send: async () => '{"tool":"read"}' });
        await started.promise;
        result.reject(error);
        await expect(outcome).rejects.toBe(error);
    });

    test.each(['mcp_list_resources', 'mcp_read_resource'])('%s rethrows cancellation', async (name) => {
        const error = new DOMException('cancelled', 'AbortError');
        const fail = jest.fn().mockRejectedValue(error);
        const client = createResourceClient([{ name: 'docs', client: { listResources: fail, readResource: fail } }]);
        await expect(client.callTool(name, { server: 'docs', uri: 'test' })).rejects.toBe(error);
    });

    test.each(['listPrompts', 'getPrompt'])('prompt import rethrows %s cancellation', async (method) => {
        const error = new DOMException('cancelled', 'AbortError');
        const client = { listPrompts: async () => [{ name: 'test' }], getPrompt: async () => ({}) };
        client[method] = jest.fn().mockRejectedValue(error);
        await expect(importServerPrompts('docs', client)).rejects.toBe(error);
    });
});
