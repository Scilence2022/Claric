/** @jest-environment jsdom */

jest.mock('../src/lib/file-store.js', () => {
    const actual = jest.requireActual('../src/lib/file-store.js');
    const store = actual.createFileStore({ adapter: actual.createMemoryFileAdapter() });
    return { ...actual, ...store };
});

const { answerQuestion } = require('../src/taskpane/word-actions.js');
const { saveFile, deleteFile } = require('../src/lib/file-store.js');
const response = (content) => ({
    ok: true, headers: { get: () => 'application/json' },
    json: async () => ({ choices: [{ message: { content }, finish_reason: 'stop' }] }),
});
const call = (tool, args) => response(JSON.stringify({ tool, args }));
const bodies = () => fetch.mock.calls.map(([, init]) => JSON.parse(init.body));

function world(fileReferences) {
    const absent = { isNullObject: true, load: jest.fn() };
    const selection = {
        text: 'Selected clause', load: jest.fn(), parentTableOrNullObject: absent,
        parentTableCellOrNullObject: absent, paragraphs: { items: [], load: jest.fn() },
    };
    global.Word = { run: jest.fn(async (fn) => fn({
        document: {
            getSelection: () => selection,
            body: { paragraphs: { items: [{ text: 'Current document body', load: jest.fn() }], load: jest.fn() } },
        }, sync: async () => {},
    })) };
    return {
        appState: {
            config: { backend: 'custom', providers: { custom: { url: 'https://fixture.invalid', model: 'fixture' } }, docExtraction: { richness: 'plain' } },
            promptManager: { getActivePrompt: () => ({ template: 'Trusted user context instruction' }) },
        },
        fileReferences, log: jest.fn(),
        conversationHistory: [
            { role: 'user', content: 'Earlier question' },
            { role: 'assistant', content: 'Earlier answer' },
        ],
    };
}

beforeEach(() => { global.fetch = jest.fn(); });
afterEach(() => { delete global.fetch; delete global.Word; });

test('real QA reads library text on demand via get/read/finish, preserving Word, selection, persona and history', async () => {
    const file = await saveFile(new File(['Library-only text. '.repeat(2000)], 'reference.txt', { type: 'text/plain' }));
    const deps = world([file]);
    fetch.mockResolvedValueOnce(call('file_get', { fileId: file.fileId }))
        .mockResolvedValueOnce(call('file_read', { fileId: file.fileId, offset: 18, limit: 18 }))
        .mockResolvedValueOnce(call('finish', { summary: 'The reference supports the selected clause.' }));
    const onToken = jest.fn();
    await expect(answerQuestion(deps, {
        question: 'What does this reference say?', selectionText: 'Selected clause',
        skillTemplate: 'Answer concisely', onToken,
    })).resolves.toBe('The reference supports the selected clause.');
    expect(fetch).toHaveBeenCalledTimes(3);
    const first = bodies()[0].messages;
    expect(first[0].content).toContain('Trusted user context instruction');
    expect(first.slice(1, 3)).toEqual(deps.conversationHistory);
    expect(first.at(-1).content).toContain('Current document body');
    expect(first.at(-1).content).toContain('Selected clause');
    expect(first.at(-1).content).toContain('Answer concisely');
    expect(first.at(-1).content).toContain(file.fileId);
    expect(JSON.stringify(bodies().slice(0, 2))).not.toContain('Library-only text.');
    const observation = JSON.parse(bodies()[2].messages.at(-1).content);
    expect(observation).toMatchObject({ ok: true, result: { text: expect.stringContaining('Library-only text'), offset: 18, limit: 18, hasMore: true } });
    expect(observation.result.text).toHaveLength(18);
    expect(onToken).toHaveBeenCalledTimes(1);
    expect(onToken).toHaveBeenCalledWith('The reference supports the selected clause.');
});

test('only submitted IDs and versions are authorized, even when history names another file', async () => {
    const file = await saveFile(new File(['Allowed'], 'allowed.txt'));
    const other = await saveFile(new File(['Private library body'], 'private.txt'));
    const deps = world([file]);
    deps.conversationHistory[0].content = `Previous file ${other.fileId}`;
    fetch.mockResolvedValueOnce(call('file_read', { fileId: other.fileId }))
        .mockResolvedValueOnce(call('file_read', { fileId: file.fileId, versionId: 'wrong' }))
        .mockResolvedValueOnce(call('delete_file', { fileId: file.fileId }))
        .mockResolvedValueOnce(call('finish', { summary: 'Those references are not authorized.' }));
    await answerQuestion(deps, { question: 'What do they say?', selectionText: 'Selected clause' });
    expect(JSON.parse(bodies()[1].messages.at(-1).content)).toMatchObject({ ok: false, error: expect.stringContaining('Access denied') });
    expect(JSON.parse(bodies()[2].messages.at(-1).content)).toMatchObject({ ok: false, error: expect.stringContaining('version') });
    expect(JSON.parse(bodies()[3].messages.at(-1).content)).toMatchObject({ ok: false, error: expect.stringContaining('Unknown tool') });
    expect(JSON.stringify(bodies())).not.toContain('Private library body');
});

test.each(['missing', 'version', 'unversioned'])('invalid attachment fails before model call: %s', async (kind) => {
    const file = await saveFile(new File(['text'], 'reference.txt'));
    if (kind === 'missing') await deleteFile(file.fileId);
    if (kind === 'version') file.versionId = 'wrong';
    if (kind === 'unversioned') delete file.versionId;
    await expect(answerQuestion(world([file]), { question: 'What?', selectionText: 'Selected clause' })).rejects.toThrow(/file|version/i);
    expect(fetch).not.toHaveBeenCalled();
});

test.each([false, true])('image reading sends real vision content and never strips it after backend rejection: reject=%s', async (reject) => {
    const file = await saveFile(new File([new Uint8Array([1, 2, 3])], 'image.png', { type: 'image/png' }));
    const deps = world([file]);
    fetch.mockResolvedValueOnce(call('file_read', { fileId: file.fileId }));
    if (reject) fetch.mockResolvedValueOnce({ ok: false, status: 400, statusText: 'Bad Request', text: async () => 'Images unsupported' });
    else fetch.mockResolvedValueOnce(call('finish', { summary: 'Image answer' }));
    const pending = answerQuestion(deps, { question: 'What is in the image?', selectionText: 'Selected clause' });
    if (reject) await expect(pending).rejects.toThrow(/HTTP 400/);
    else await expect(pending).resolves.toBe('Image answer');
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(bodies()[1].messages.at(-1).content).toContainEqual({ type: 'image_url', image_url: { url: 'data:image/png;base64,AQID' } });
});

test('abort between model response and file access propagates without a final answer', async () => {
    const file = await saveFile(new File(['text'], 'reference.txt'));
    const controller = new AbortController();
    fetch.mockImplementationOnce(async () => {
        controller.abort();
        return call('file_read', { fileId: file.fileId });
    });
    const onToken = jest.fn();
    await expect(answerQuestion(world([file]), {
        question: 'What?', selectionText: 'Selected clause', signal: controller.signal, onToken,
    })).rejects.toMatchObject({ name: 'AbortError' });
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(onToken).not.toHaveBeenCalled();
});

test('first-turn file QA preserves trusted context and temporary image inputs throughout the loop', async () => {
    const file = await saveFile(new File(['text'], 'reference.txt'));
    const deps = world([file]);
    deps.conversationHistory = [];
    fetch.mockResolvedValueOnce(call('file_get', { fileId: file.fileId }))
        .mockResolvedValueOnce(call('finish', { summary: 'Answer from the supplied inputs' }));
    await answerQuestion(deps, {
        question: 'What?', selectionText: 'Selected clause', questionImages: [{ dataUrl: 'data:image/png;base64,AQID' }],
    });
    for (const body of bodies()) {
        expect(body.messages[0].content).toContain('Trusted user context instruction');
        expect(body.messages[1].content).toContainEqual({ type: 'image_url', image_url: { url: 'data:image/png;base64,AQID' } });
    }
});

test('a file deleted during the loop is an error observation, not stale cached content', async () => {
    const file = await saveFile(new File(['deleted body'], 'reference.txt'));
    fetch.mockImplementationOnce(async () => {
        await deleteFile(file.fileId);
        return call('file_read', { fileId: file.fileId });
    }).mockResolvedValueOnce(call('finish', { summary: 'The file was deleted; please attach it again.' }));
    await answerQuestion(world([file]), { question: 'What?', selectionText: 'Selected clause' });
    expect(JSON.parse(bodies()[1].messages.at(-1).content).ok).toBe(false);
    expect(JSON.stringify(bodies())).not.toContain('deleted body');
});

test('pre-aborted QA makes no Word or model calls', async () => {
    const deps = world([]);
    const controller = new AbortController();
    controller.abort();
    await expect(answerQuestion(deps, { question: 'What?', signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(Word.run).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
});

test('temporary-image-only QA retains its original 4xx text fallback', async () => {
    fetch.mockResolvedValueOnce({ ok: false, status: 400, statusText: 'Bad Request', text: async () => 'Images unsupported' })
        .mockResolvedValueOnce(response('Text fallback answer'));
    await expect(answerQuestion(world([]), {
        question: 'What?', selectionText: 'Selected clause', questionImages: [{ dataUrl: 'data:image/png;base64,AQID' }],
    })).resolves.toBe('Text fallback answer');
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(typeof bodies()[1].messages.at(-1).content).toBe('string');
});

test('real list/search/get/read loop scopes metadata to current references and never eagerly loads text', async () => {
    const reference = await saveFile(new File(['Selected reference body'], 'chosen-reference.txt'));
    const other = await saveFile(new File(['Unrelated private body'], 'other-reference.txt'));
    fetch.mockResolvedValueOnce(call('file_list', { allowedIds: [other.fileId], offset: 0, limit: 10 }))
        .mockResolvedValueOnce(call('file_search', { query: 'reference', offset: 0, limit: 10 }))
        .mockResolvedValueOnce(call('file_get', { fileId: reference.fileId }))
        .mockResolvedValueOnce(call('file_read', { fileId: reference.fileId, offset: 0, limit: 8 }))
        .mockResolvedValueOnce(call('finish', { summary: 'Selected reference answer' }));
    await expect(answerQuestion(world([reference]), { question: 'What is in the reference?', selectionText: 'Selected clause' }))
        .resolves.toBe('Selected reference answer');
    const requests = bodies();
    for (const index of [1, 2]) {
        const observation = JSON.parse(requests[index].messages.at(-1).content);
        expect(observation.ok).toBe(true);
        expect(observation.result.total).toBe(1);
        expect(observation.result.items.map((item) => item.fileId)).toEqual([reference.fileId]);
        expect(observation.result.items[0].versionId).toBe(reference.versionId);
    }
    expect(JSON.stringify(requests.slice(0, 4))).not.toContain('Selected reference body');
    expect(JSON.stringify(requests.flatMap((request) => request.messages.filter((message) => message.role !== 'assistant')))).not.toContain(other.fileId);
    expect(JSON.stringify(requests)).not.toContain('Unrelated private body');
    expect(JSON.parse(requests[4].messages.at(-1).content)).toMatchObject({ ok: true, result: { text: 'Selected' } });
});

test.each(['file_list', 'file_search'])('%s checks every authorized reference before listing metadata', async (tool) => {
    const first = await saveFile(new File(['first'], 'first.txt'));
    const second = await saveFile(new File(['second'], 'second.txt'));
    fetch.mockImplementationOnce(async () => {
        await deleteFile(second.fileId);
        return call(tool, { query: 'first' });
    }).mockResolvedValueOnce(call('finish', { summary: 'An attached file was deleted.' }));
    await answerQuestion(world([first, second]), { question: 'What files are attached?', selectionText: 'Selected clause' });
    expect(JSON.parse(bodies()[1].messages.at(-1).content)).toMatchObject({ ok: false, error: expect.stringContaining('not found') });
});
