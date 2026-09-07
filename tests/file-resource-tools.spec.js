const { createFileStore, createMemoryFileAdapter } = require('../src/lib/file-store.js');
const { createFileResourceToolExecutor, executeFileResourceTool, FILE_RESOURCE_TOOL_SPECS } = require('../src/lib/file-resource-tools.js');
const { runToolLoop } = require('../src/lib/tool-loop.js');

it('requires explicit scope, denies arbitrary IDs and ignores caller scope expansion', async () => {
    expect(() => createFileResourceToolExecutor()).toThrow('allowedIds');
    expect(await executeFileResourceTool('file_list')).toMatchObject({ ok: false });
    const store = createFileStore({ adapter: createMemoryFileAdapter() });
    const shared = await store.saveFile(new File(['shared text'], 'shared.txt'));
    const secret = await store.saveFile(new File(['private text'], 'private.txt'));
    const ids = [shared.fileId];
    const execute = createFileResourceToolExecutor({ store, allowedIds: ids });
    ids.push(secret.fileId);
    const list = await execute('file_list', { allowedIds: [secret.fileId] });
    expect(list.result.items.map((item) => item.fileId)).toEqual([shared.fileId]);
    for (const tool of ['file_get', 'file_read']) {
        expect(await execute(tool, { fileId: secret.fileId })).toMatchObject({ ok: false, error: expect.stringContaining('not shared') });
        expect(await execute(tool, { fileId: '/etc/passwd' })).toMatchObject({ ok: false });
    }
    expect(await execute('file_read', { fileId: shared.fileId, limit: 6 })).toMatchObject({ ok: true, result: { text: 'shared', hasMore: true } });
    expect(await execute('file_get', { fileId: shared.fileId })).toMatchObject({ ok: true, result: { fileId: shared.fileId } });
    expect(await execute('file_search', { query: 'PRIVATE' })).toMatchObject({ result: { total: 0 } });
    expect(await execute('deleteFile', { fileId: shared.fileId })).toMatchObject({ ok: false });
    expect(await execute('file_list', [])).toMatchObject({ ok: false });
    await store.deleteFile(shared.fileId);
    expect(await execute('file_read', { fileId: shared.fileId })).toMatchObject({ ok: false, error: expect.stringContaining('deleted') });
});

it('returns image attachments through the actual loop multimodal wire channel, not result JSON', async () => {
    const store = createFileStore({ adapter: createMemoryFileAdapter() });
    const image = await store.saveFile(new File([new Uint8Array([137, 80, 78, 71])], 'pic.png'));
    const execute = createFileResourceToolExecutor({ store, allowedIds: [image.fileId] });
    const sent = [];
    await runToolLoop({
        systemPrompt: 'Read only', taskPrompt: 'Inspect image', tools: FILE_RESOURCE_TOOL_SPECS, execute,
        send: async (messages) => {
            sent.push(JSON.parse(JSON.stringify(messages)));
            return sent.length === 1 ? JSON.stringify({ tool: 'file_read', args: { fileId: image.fileId } }) : '{"tool":"finish","args":{"summary":"Inspected"}}';
        },
    });
    const content = sent[1].at(-1).content;
    expect(content[1]).toEqual({ type: 'image_url', image_url: { url: 'data:image/png;base64,iVBORw==' } });
    expect(content[0].text).not.toContain('base64');
    expect(JSON.parse(content[0].text)).toMatchObject({ ok: true, result: { file: { fileId: image.fileId } } });
});

it('propagates cancellation before and after a tool call', async () => {
    const controller = new AbortController();
    const store = { listFiles: async () => { controller.abort(); return { items: [], total: 0 }; } };
    const execute = createFileResourceToolExecutor({ store, allowedIds: [], signal: controller.signal });
    await expect(execute('file_list')).rejects.toMatchObject({ name: 'AbortError' });
    await expect(execute('file_list')).rejects.toMatchObject({ name: 'AbortError' });
});
