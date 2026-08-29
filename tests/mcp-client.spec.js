/**
 * MCP client tests (Streamable HTTP JSON-RPC, injected fetch).
 *
 * Covers: initialize handshake + session-id echo, token auth header,
 * notifications/initialized, tools/list, tools/call, JSON response bodies
 * AND SSE-encoded response bodies (Streamable HTTP allows both), JSON-RPC
 * error propagation, and HTTP error propagation.
 */

const { connectMcpServer } = require('../src/lib/mcp-client.js');

/** Records POSTs and replies from a scripted queue of responses. */
function makeFetchMock(responses) {
    const posts = [];
    const fetchFn = async (url, options = {}) => {
        posts.push({ url, options });
        const next = responses.shift();
        if (!next) throw new Error('unexpected extra fetch call');
        return typeof next === 'function' ? next(posts.length, posts) : next;
    };
    return { fetchFn, posts };
}

function jsonResponse(body, headers = {}, status = 200) {
    return {
        ok: status >= 200 && status < 300,
        status,
        headers: { get: (n) => headers[n.toLowerCase()] ?? null },
        text: async () => JSON.stringify(body),
    };
}

const INIT_RESULT = {
    jsonrpc: '2.0',
    id: 1,
    result: { protocolVersion: '2025-06-18', serverInfo: { name: 'mock-mcp', version: '0.1.0' } },
};

function sseResponse(dataLines, headers = {}, status = 200) {
    const body = dataLines.map((d) => `data: ${JSON.stringify(d)}\n\n`).join('');
    return {
        ok: status >= 200 && status < 300,
        status,
        headers: { get: (n) => headers[n.toLowerCase()] ?? (n === 'content-type' ? 'text/event-stream' : null) },
        text: async () => body,
    };
}

describe('connectMcpServer', () => {
    test('handshake: initialize, session-id echo, initialized notification', async () => {
        const { fetchFn, posts } = makeFetchMock([
            jsonResponse(INIT_RESULT, { 'mcp-session-id': 'sess-123' }),
            { ok: true, status: 202, headers: { get: () => null }, text: async () => '' },
            jsonResponse({ jsonrpc: '2.0', id: 2, result: { tools: [] } }),
        ]);

        const client = await connectMcpServer({ url: 'https://mcp.example/mcp', token: 'secret', fetchFn });
        expect(client.serverInfo).toEqual({ name: 'mock-mcp', version: '0.1.0' });

        // Initialize carried protocol + client info + auth.
        const init = JSON.parse(posts[0].options.body);
        expect(init.method).toBe('initialize');
        expect(init.params.protocolVersion).toBe('2025-06-18');
        expect(init.params.clientInfo.name).toBe('claric');
        expect(posts[0].options.headers['Authorization']).toBe('Bearer secret');

        // Session id from the response header is echoed on later requests.
        const notif = JSON.parse(posts[1].options.body);
        expect(notif.method).toBe('notifications/initialized');
        expect(posts[1].options.headers['Mcp-Session-Id']).toBe('sess-123');

        await client.listTools();
        expect(JSON.parse(posts[2].options.body).method).toBe('tools/list');
        expect(posts[2].options.headers['Mcp-Session-Id']).toBe('sess-123');
    });

    test('tools/list returns the tools array', async () => {
        const tools = [{ name: 'web_search', description: 'Search the web', inputSchema: { type: 'object' } }];
        const { fetchFn } = makeFetchMock([
            jsonResponse(INIT_RESULT),
            { ok: true, status: 202, headers: { get: () => null }, text: async () => '' },
            jsonResponse({ jsonrpc: '2.0', id: 2, result: { tools } }),
        ]);
        const client = await connectMcpServer({ url: 'https://mcp.example/mcp', fetchFn });
        expect(await client.listTools()).toEqual(tools);
    });

    test('tools/call parses a plain JSON response', async () => {
        const { fetchFn } = makeFetchMock([
            jsonResponse(INIT_RESULT),
            { ok: true, status: 202, headers: { get: () => null }, text: async () => '' },
            jsonResponse({ jsonrpc: '2.0', id: 2, result: { content: [{ type: 'text', text: '42' }], isError: false } }),
        ]);
        const client = await connectMcpServer({ url: 'https://mcp.example/mcp', fetchFn });
        const result = await client.callTool('calc', { x: 1 });
        expect(result.content[0].text).toBe('42');
    });

    test('tools/call handles an SSE-encoded response body', async () => {
        // Replies must echo the request id; the module-level counter keeps
        // advancing across tests, so derive the id from the captured POST.
        const { fetchFn } = makeFetchMock([
            jsonResponse(INIT_RESULT),
            { ok: true, status: 202, headers: { get: () => null }, text: async () => '' },
            (postIndex, posts) => {
                const id = JSON.parse(posts[postIndex - 1].options.body).id;
                return sseResponse([
                    { jsonrpc: '2.0', method: 'notifications/progress', params: { value: 1 } },
                    { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: 'streamed' }] } },
                ]);
            },
        ]);
        const client = await connectMcpServer({ url: 'https://mcp.example/mcp', fetchFn });
        const result = await client.callTool('search', { q: 'x' });
        expect(result.content[0].text).toBe('streamed');
    });

    test('JSON-RPC error results throw with code and message', async () => {
        const { fetchFn } = makeFetchMock([
            jsonResponse(INIT_RESULT),
            { ok: true, status: 202, headers: { get: () => null }, text: async () => '' },
            jsonResponse({ jsonrpc: '2.0', id: 2, error: { code: -32601, message: 'Method not found' } }),
        ]);
        const client = await connectMcpServer({ url: 'https://mcp.example/mcp', fetchFn });
        await expect(client.listTools()).rejects.toThrow(/-32601.*Method not found/);
    });

    test('HTTP errors propagate with the status code', async () => {
        const { fetchFn } = makeFetchMock([
            jsonResponse({ message: 'unauthorized' }, {}, 401),
        ]);
        await expect(connectMcpServer({ url: 'https://mcp.example/mcp', fetchFn }))
            .rejects.toThrow(/401/);
    });
});

describe('prompts and resources helpers', () => {
    test('listPrompts/getPrompt/listResources/readResource parse results and default empty', async () => {
        const { fetchFn } = makeFetchMock([
            jsonResponse(INIT_RESULT),
            { ok: true, status: 202, headers: { get: () => null }, text: async () => '' },
            jsonResponse({ jsonrpc: '2.0', id: 2, result: { prompts: [{ name: 'review', description: 'Review a draft' }] } }),
            jsonResponse({ jsonrpc: '2.0', id: 3, result: { messages: [{ role: 'user', content: { type: 'text', text: 'Review: {topic}' } }] } }),
            jsonResponse({ jsonrpc: '2.0', id: 4, result: { resources: [{ uri: 'file:///a.md', name: 'a' }] } }),
            jsonResponse({ jsonrpc: '2.0', id: 5, result: { contents: [{ uri: 'file:///a.md', text: 'content' }] } }),
            // Unknown methods (server without prompts/resources) degrade to empty.
            jsonResponse({ jsonrpc: '2.0', id: 6, error: { code: -32601, message: 'Method not found' } }),
        ]);
        const client = await connectMcpServer({ url: 'https://mcp.example/mcp', fetchFn });

        expect(await client.listPrompts()).toEqual([{ name: 'review', description: 'Review a draft' }]);
        const prompt = await client.getPrompt('review', { topic: 'x' });
        expect(prompt.messages[0].content.text).toContain('{topic}');
        expect(await client.listResources()).toEqual([{ uri: 'file:///a.md', name: 'a' }]);
        expect((await client.readResource('file:///a.md')).contents[0].text).toBe('content');
        // A server without resource support degrades to an empty list, not a crash.
        await expect(client.listResources()).resolves.toEqual([]);
    });
});
