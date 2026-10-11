/**
 * MCP tools bridge tests: name sanitization/namespacing, schema→example
 * derivation, executor routing, error/isError/image handling, truncation.
 */

const {
    buildLoopTools,
    createMcpToolExecutor,
    createResourceClient,
    importServerPrompts,
    RESOURCE_TOOL_SPECS,
    MAX_MCP_RESULT_CHARS,
    MCP_APPROVAL,
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
            { name: 's', client, mcpTools: [{ name: 'greet', description: '', inputSchema: {},
                annotations: { readOnlyHint: true } }] },
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

describe('MCP tool-call approval', () => {
    function setup(confirmCall, annotations) {
        const client = makeClient({});
        const { mapping } = buildLoopTools([
            { name: 'mail', client, mcpTools: [{ name: 'send_email', description: '', inputSchema: {}, annotations }] },
        ]);
        return { client, execute: createMcpToolExecutor(mapping, { confirmCall }) };
    }

    test('records the server-declared read-only hint in the mapping', () => {
        const { mapping } = buildLoopTools([{ name: 's', client: makeClient({}), mcpTools: [
            { name: 'search', annotations: { readOnlyHint: true } },
            { name: 'write', annotations: { readOnlyHint: false } },
            { name: 'plain' },
        ] }]);
        expect([...mapping.values()].map((entry) => entry.readOnly)).toEqual([true, false, false]);
    });

    test('a tool not declared read-only is refused when no approver is available', async () => {
        const { client, execute } = setup(undefined);
        const observation = await execute('send_email', { to: 'x@example.com' });
        expect(observation).toEqual({ ok: false, error: expect.stringContaining("needs the user's approval") });
        expect(client.callTool).not.toHaveBeenCalled();
    });

    test('the approver sees the server, original tool name and arguments', async () => {
        const confirm = jest.fn(async () => MCP_APPROVAL.ONCE);
        const { client, execute } = setup(confirm);
        await execute('send_email', { to: 'x@example.com' });
        expect(confirm).toHaveBeenCalledWith({ toolName: 'send_email', tool: 'send_email', serverName: 'mail',
            args: { to: 'x@example.com' } });
        expect(client.callTool).toHaveBeenCalledWith('send_email', { to: 'x@example.com' });
    });

    test('a declined call never reaches the server and tells the model not to retry', async () => {
        const { client, execute } = setup(jest.fn(async () => MCP_APPROVAL.DENY));
        const observation = await execute('send_email', {});
        expect(observation).toEqual({ ok: false, error: expect.stringContaining('declined') });
        expect(observation.error).toContain('Do not retry');
        expect(client.callTool).not.toHaveBeenCalled();
    });

    test('an unexpected approver answer counts as a refusal', async () => {
        const { client, execute } = setup(jest.fn(async () => true));
        expect((await execute('send_email', {})).ok).toBe(false);
        expect(client.callTool).not.toHaveBeenCalled();
    });

    test('"once" asks again on the next call; "run" approves the tool for the rest of the run', async () => {
        const once = jest.fn(async () => MCP_APPROVAL.ONCE);
        const first = setup(once);
        await first.execute('send_email', {});
        await first.execute('send_email', {});
        expect(once).toHaveBeenCalledTimes(2);

        const run = jest.fn(async () => MCP_APPROVAL.RUN);
        const second = setup(run);
        await second.execute('send_email', {});
        await second.execute('send_email', {});
        expect(run).toHaveBeenCalledTimes(1);
        expect(second.client.callTool).toHaveBeenCalledTimes(2);
    });

    test('read-only tools and the synthetic resource tools run without asking', async () => {
        const confirm = jest.fn();
        const readOnly = setup(confirm, { readOnlyHint: true });
        await readOnly.execute('send_email', {});
        const docs = makeClient({});
        docs.listResources = async () => [];
        const { mapping } = buildLoopTools([
            { name: 'resources', client: createResourceClient([{ name: 'docs', client: docs }]), mcpTools: RESOURCE_TOOL_SPECS },
        ]);
        await createMcpToolExecutor(mapping, { confirmCall: confirm })('mcp_list_resources', {});
        expect(confirm).not.toHaveBeenCalled();
        expect(readOnly.client.callTool).toHaveBeenCalledTimes(1);
    });

    test('cancelling while approval is pending propagates the AbortError', async () => {
        const error = new DOMException('cancelled', 'AbortError');
        const { client, execute } = setup(jest.fn(async () => { throw error; }));
        await expect(execute('send_email', {})).rejects.toBe(error);
        expect(client.callTool).not.toHaveBeenCalled();
    });
});

describe('resource tools and prompt convergence (phase C)', () => {
    test('resource client lists across servers and reads by server+uri', async () => {
        const docs = makeClient({});
        docs.listResources = async () => [{ uri: 'file:///a.md', name: 'a' }];
        docs.readResource = async (uri) => ({ contents: [{ uri, text: `content of ${uri}` }] });
        const web = makeClient({});
        web.listResources = async () => { throw new Error('Method not found'); };

        const resourceClient = createResourceClient([
            { name: 'docs', client: docs },
            { name: 'web', client: web },
        ]);
        const { loopTools, mapping } = buildLoopTools([
            { name: 'resources', client: resourceClient, mcpTools: RESOURCE_TOOL_SPECS },
        ]);
        const execute = createMcpToolExecutor(mapping);

        expect(loopTools.map((t) => t.name)).toEqual(['mcp_list_resources', 'mcp_read_resource']);
        const listing = await execute('mcp_list_resources', {});
        expect(listing.result).toBe('docs | file:///a.md | a');
        const reading = await execute('mcp_read_resource', { server: 'docs', uri: 'file:///a.md' });
        expect(reading.result).toBe('content of file:///a.md');
        const badServer = await execute('mcp_read_resource', { server: 'nope', uri: 'x' });
        expect(badServer.ok).toBe(false);
    });

    test('importServerPrompts converts MCP prompts into skill descriptors', async () => {
        const client = makeClient({});
        client.listPrompts = async () => [
            { name: 'review-draft', description: 'Review a draft' },
            { name: 'empty-one', description: 'no body' },
            { name: 'Broken!', description: 'bad' },
        ];
        client.getPrompt = async (name) => {
            if (name === 'review-draft') {
                return { messages: [{ role: 'user', content: { type: 'text', text: 'Review {selection} for gaps.' } }] };
            }
            if (name === 'empty-one') return { messages: [] };
            throw new Error('bad prompt');
        };

        const { imported, errors } = await importServerPrompts('paper', client);
        expect(imported).toHaveLength(1);
        expect(imported[0]).toEqual(expect.objectContaining({
            name: 'paper-review-draft',
            slash: '/paper-review-draft',
            category: 'chat',
            defaultTemplate: 'Review {selection} for gaps.',
        }));
        expect(errors).toHaveLength(2);
    });

    test('importServerPrompts degrades when the server has no prompts', async () => {
        const client = makeClient({});
        client.listPrompts = async () => { throw new Error('Method not found'); };
        const { imported, errors } = await importServerPrompts('s', client);
        expect(imported).toEqual([]);
        expect(errors).toHaveLength(1);
    });
});
