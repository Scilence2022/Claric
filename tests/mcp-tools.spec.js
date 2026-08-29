/**
 * MCP tools bridge tests: name sanitization/namespacing, schema→example
 * derivation, executor routing, error/isError/image handling, truncation.
 */

const {
    buildLoopTools,
    createMcpToolExecutor,
    MAX_MCP_RESULT_CHARS,
} = require('../src/lib/mcp-tools.js');
const { defineTool } = require('../src/lib/tool-registry.js');

function makeClient(resultsByName) {
    return {
        callTool: jest.fn(async (name) => {
            if (resultsByName[name] instanceof Error) throw resultsByName[name];
            return resultsByName[name] ?? { content: [{ type: 'text', text: `ok:${name}` }] };
        }),
    };
}

describe('buildLoopTools', () => {
    test('sanitizes MCP names into the loop namespace and maps them back', () => {
        const client = makeClient({});
        const { loopTools, mapping } = buildLoopTools([
            {
                name: 'docs',
                client,
                mcpTools: [
                    { name: 'web-search', description: 'Search the web', inputSchema: { type: 'object', properties: { query: { type: 'string' } } } },
                    { name: 'read_file', description: 'Read a file', inputSchema: { type: 'object' } },
                    { name: '9lives', description: 'Starts with a digit', inputSchema: {} },
                ],
            },
        ]);

        expect(loopTools.map((t) => t.name)).toEqual(['web_search', 'read_file', 'mcp_9lives']);
        expect(mapping.get('web_search')).toEqual(expect.objectContaining({ originalName: 'web-search', serverName: 'docs' }));
        // All produced specs pass the loop's own validation.
        for (const spec of loopTools) expect(() => defineTool(spec)).not.toThrow();
    });

    test('namespaces collisions across servers instead of misrouting', () => {
        const { mapping } = buildLoopTools([
            { name: 'serverA', client: makeClient({}), mcpTools: [{ name: 'search', description: '', inputSchema: {} }] },
            { name: 'serverB', client: makeClient({}), mcpTools: [{ name: 'search', description: '', inputSchema: {} }] },
        ]);
        expect([...mapping.keys()].sort()).toEqual(['search', 'search_serverb']);
        expect(mapping.get('search').serverName).toBe('serverA');
        expect(mapping.get('search_serverb').serverName).toBe('serverB');
    });

    test('derives example args from the JSON Schema', () => {
        const { loopTools } = buildLoopTools([
            {
                name: 's',
                client: makeClient({}),
                mcpTools: [{
                    name: 'doit',
                    description: '',
                    inputSchema: {
                        type: 'object',
                        properties: {
                            q: { type: 'string' },
                            n: { type: 'integer' },
                            flag: { type: 'boolean' },
                            tags: { type: 'array' },
                            meta: { type: 'object' },
                        },
                    },
                }],
            },
        ]);
        expect(loopTools[0].argsExample).toEqual({ q: '...', n: 0, flag: true, tags: [], meta: {} });
    });

    test('descriptions carry the server of origin', () => {
        const { loopTools } = buildLoopTools([
            { name: 'math', client: makeClient({}), mcpTools: [{ name: 'add', description: 'Add two numbers', inputSchema: {} }] },
        ]);
        expect(loopTools[0].description).toBe('[MCP:math] Add two numbers');
    });
});

describe('createMcpToolExecutor', () => {
    function setup(resultsByName) {
        const client = makeClient(resultsByName);
        const { mapping } = buildLoopTools([
            { name: 's', client, mcpTools: [{ name: 'greet', description: '', inputSchema: {} }] },
        ]);
        return { client, execute: createMcpToolExecutor(mapping) };
    }

    test('routes a loop call to the original MCP tool name and returns text', async () => {
        const { client, execute } = setup({ greet: { content: [{ type: 'text', text: 'hello' }] } });
        const observation = await execute('greet', { who: 'world' });
        expect(observation).toEqual({ ok: true, result: 'hello' });
        expect(client.callTool).toHaveBeenCalledWith('greet', { who: 'world' });
    });

    test('joins multiple text blocks', async () => {
        const { execute } = setup({ greet: { content: [{ type: 'text', text: 'a' }, { type: 'text', text: 'b' }] } });
        expect((await execute('greet', {})).result).toBe('a\nb');
    });

    test('image content blocks become loop attachments', async () => {
        const { execute } = setup({
            greet: { content: [{ type: 'text', text: 'see attached' }, { type: 'image', data: 'QUJD', mimeType: 'image/jpeg' }] },
        });
        const observation = await execute('greet', {});
        expect(observation.attachments).toEqual([{ dataUrl: 'data:image/jpeg;base64,QUJD' }]);
        expect(observation.result).toBe('see attached');
    });

    test('isError results become error observations', async () => {
        const { execute } = setup({ greet: { isError: true, content: [{ type: 'text', text: 'boom' }] } });
        expect(await execute('greet', {})).toEqual({ ok: false, error: 'boom' });
    });

    test('transport failures become error observations (loop can self-correct)', async () => {
        const { execute } = setup({ greet: new Error('socket hung up') });
        expect(await execute('greet', {})).toEqual({ ok: false, error: expect.stringContaining('socket hung up') });
    });

    test('unknown loop tool names are error observations', async () => {
        const { execute } = setup({});
        expect(await execute('nope', {})).toEqual({ ok: false, error: expect.stringContaining('Unknown MCP tool') });
    });

    test('oversized results are truncated, never dropped', async () => {
        const big = 'x'.repeat(MAX_MCP_RESULT_CHARS + 1000);
        const { execute } = setup({ greet: { content: [{ type: 'text', text: big }] } });
        const observation = await execute('greet', {});
        expect(observation.ok).toBe(true);
        expect(observation.result.length).toBeLessThan(big.length);
        expect(observation.result).toMatch(/\[truncated/);
    });
});
