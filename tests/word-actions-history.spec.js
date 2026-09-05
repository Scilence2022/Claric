/** @jest-environment jsdom */

jest.mock('../src/lib/comment-extractor.js', () => ({
    extractDocumentStructured: jest.fn(async () => 'Current document'),
    extractAllComments: jest.fn(async () => []),
    extractCommentsOnRange: jest.fn(async () => []),
    estimateTokenCount: jest.fn(() => 4),
}));
jest.mock('../src/lib/document-generator.js', () => ({
    buildSummaryHtml: jest.fn((text) => text),
    createSummaryDocument: jest.fn(async () => {}),
}));
jest.mock('../src/lib/document-parser.js', () => ({ parseDocument: jest.fn() }));
jest.mock('../src/lib/document-chunker.js', () => ({ chunkDocument: jest.fn() }));
jest.mock('../src/lib/context-extractor.js', () => ({
    extractContext: jest.fn(() => ({ definitions: [], outline: [] })),
    formatContextPrefix: jest.fn(() => 'Document context'),
}));
jest.mock('../src/lib/reassembler.js', () => ({
    bookmarkChunkRanges: jest.fn(async () => new Map([['one', '_one'], ['two', '_two']])),
    cleanupBookmarks: jest.fn(async () => {}),
    applyChunkResults: jest.fn(async () => ({ amendmentsApplied: 0, commentsInserted: 0 })),
}));

const actions = require('../src/taskpane/word-actions.js');
const { parseDocument } = require('../src/lib/document-parser.js');
const { chunkDocument } = require('../src/lib/document-chunker.js');
const { resumeCommentFromBookmark } = require('../src/lib/comment-request.js');
const { sendPrompt } = require('../src/lib/llm-client.js');

const history = [
    { role: 'user', content: 'Use the second proposal and keep its formal tone.' },
    { role: 'assistant', content: 'The second proposal is a concise payment clause.' },
];
const config = { url: 'https://fixture.invalid', model: 'fixture-model', apiKey: '', apiPath: '/v1' };
const system = { role: 'system', content: 'Trusted context instruction' };
const svg = '<svg width="100" height="50" viewBox="0 0 100 50"><rect width="100" height="50"/></svg>';
const chunk = (id) => ({ id, paragraphs: [{ text: `Paragraph ${id}` }], tokenCount: 5 });
const response = (content) => ({
    ok: true,
    headers: { get: () => 'application/json' },
    json: async () => ({ choices: [{ message: { content }, finish_reason: 'stop' }] }),
});
const failure = () => ({ ok: false, status: 400, statusText: 'Bad Request', text: async () => 'Images unsupported' });
const bodies = () => fetch.mock.calls.map(([, init]) => JSON.parse(init.body));
const settle = async () => { for (let i = 0; i < 30; i++) await Promise.resolve(); };

function world() {
    const absent = { isNullObject: true, load: jest.fn() };
    const selection = {
        text: 'Selected clause', load: jest.fn(), insertBookmark: jest.fn(),
        getOoxml: () => ({ value: '<baseline/>' }),
        parentTableOrNullObject: absent, parentTableCellOrNullObject: absent,
        paragraphs: { items: [], load: jest.fn() },
    };
    const context = {
        document: {
            getSelection: () => selection,
            getBookmarkRangeOrNullObject: () => selection,
            deleteBookmark: jest.fn(),
            body: { getRange: () => selection },
            properties: { title: 'Fixture', load: jest.fn() },
        },
        sync: async () => {},
    };
    global.Word = {
        run: jest.fn(async (fn) => fn(context)),
        RangeLocation: { start: 'Start', end: 'End' },
    };
    const compose = (text) => [{ ...system }, { role: 'user', content: text }];
    const deps = {
        appState: {
            config: { backend: 'custom', providers: { custom: config } },
            promptManager: {
                getActivePrompt: (category) => category === 'context' ? { template: system.content } : null,
                composeMessages: jest.fn(compose),
                composeMergedMessages: jest.fn(compose),
                composeSummaryMessages: jest.fn(() => compose('Summarize the current document')),
            },
        },
        conversationHistory: history.map((m) => ({ ...m })),
        log: jest.fn(), logWithRetry: jest.fn(), updateStatusBar: jest.fn(),
    };
    return { deps, selection };
}

function expectHistory(messages, hasSystem = false) {
    expect(messages.map((m) => m.role)).toEqual(hasSystem
        ? ['system', 'user', 'assistant', 'user'] : ['user', 'assistant', 'user']);
    if (hasSystem) expect(messages[0].content).toContain(system.content);
    expect(messages.slice(hasSystem ? 1 : 0, hasSystem ? 3 : 2)).toEqual(history);
}

beforeEach(() => {
    jest.clearAllMocks();
    global.fetch = jest.fn(async () => response('Answer'));
    parseDocument.mockResolvedValue({ paragraphs: [], totalTokens: 10 });
    chunkDocument.mockReturnValue([chunk('one'), chunk('two')]);
});
afterEach(() => { delete global.fetch; delete global.Word; });

test.each([false, true])('Q&A preserves roles and uploaded image fallback, uploaded=%s', async (uploaded) => {
    const { deps } = world();
    if (uploaded) fetch.mockResolvedValueOnce(failure());
    const onToken = jest.fn();
    const answer = await actions.answerQuestion(deps, {
        question: 'Make that shorter', selectionText: 'Selected clause', onToken,
        questionImages: uploaded ? [{ dataUrl: 'data:image/png;base64,aW1hZ2U=' }] : [],
    });
    expect(answer).toBe('Answer');
    expect(onToken).toHaveBeenCalledWith('Answer');
    expect(fetch).toHaveBeenCalledTimes(uploaded ? 2 : 1);
    for (const body of bodies()) expectHistory(body.messages, true);
    const last = bodies().at(-1).messages.at(-1).content;
    expect(last).toContain('Make that shorter');
    expect(last).toContain('Current document');
    expect(last).not.toContain(system.content);
    if (uploaded) expect(bodies()[0].messages.at(-1).content[1].type).toBe('image_url');
});

test.each([false, true])('selection and summary use role-preserving transport, stream=%s', async (stream) => {
    const { deps } = world();
    const args = stream ? { onToken: jest.fn() } : {};
    const proposal = await actions.prepareSelectionAmendment(deps, { promptTemplate: 'Revise', ...args });
    expect(proposal.amendedText).toBe('Answer');
    await actions.runSummarySkill(deps, { promptTemplate: 'Summarize', ...args });
    for (const body of bodies()) {
        expectHistory(body.messages, true);
        expect(body.stream).toBe(stream);
    }
});

test('merged-selection classification fallback retains history exactly once', async () => {
    const { deps } = world();
    fetch.mockResolvedValueOnce(response('Undelimited response'))
        .mockResolvedValueOnce(response('===AMENDMENT===\nRevised clause\n===COMMENT===\nReview'));
    const proposal = await actions.prepareSelectionAmendment(deps, { promptTemplate: 'Revise', commentInstructions: 'Review' });
    expect(proposal.amendedText).toBe('Revised clause');
    expect(fetch).toHaveBeenCalledTimes(2);
    for (const body of bodies()) {
        expect(body.messages.filter((m) => m.role === 'assistant')).toEqual([history[1]]);
        expect(body.messages.at(-1).role).toBe('user');
    }
});

test.each([false, true])('append, table creation, plan, format and illustration requests retain history, stream=%s', async (stream) => {
    const { deps } = world();
    const args = stream ? { onToken: jest.fn() } : {};
    const cases = [
        ['prepareDocumentAppend', { instruction: 'Continue that clause' }, 'New paragraph'],
        ['prepareTableProposal', { instruction: 'Create a table with that content' }, '{"rows":[["A","B"]],"position":"end"}'],
        ['planDocumentTasks', { instruction: 'Apply that and add a heading' }, '[{"type":"edit","instruction":"Revise"}]'],
        ['prepareFormatProposal', { instruction: 'Use that formatting' }, '[{"font":{"bold":true}}]'],
        ['prepareIllustrationProposal', { instruction: 'Draw SVG at document end' }, svg],
    ];
    for (const [name, options, output] of cases) {
        fetch.mockResolvedValueOnce(response(output));
        const result = await actions[name](deps, { ...options, ...args });
        expect(result).toBeTruthy();
        const body = bodies().at(-1);
        expectHistory(body.messages);
        expect(body.messages.at(-1).content).toContain(options.instruction);
        expect(body.stream).toBe(stream);
    }
});

test('multi-cell table patch requests preserve system and previous assistant roles', async () => {
    const { deps, selection } = world();
    selection.parentTableOrNullObject = {
        isNullObject: false, rowCount: 1, values: [['A', 'B']], isUniform: true, load: jest.fn(),
    };
    selection.getRange = (loc) => ({ parentTableCellOrNullObject: {
        isNullObject: false, rowIndex: 0, cellIndex: loc === 'Start' ? 0 : 1, load: jest.fn(),
    } });
    fetch.mockResolvedValueOnce(response('{"edits":[{"row":1,"col":1,"text":"Revised"}]}'));
    const proposal = await actions.prepareSelectionAmendment(deps, { promptTemplate: 'Revise using that approach' });
    expect(proposal.tablePatch).toBeDefined();
    expectHistory(bodies()[0].messages, true);
    expect(bodies()[0].messages.at(-1).content).toContain('Revise using that approach');
});

test.each([false, true])('image resolver retains long-answer referents and SVG fallback history, fallback=%s', async (fallback) => {
    const { deps } = world();
    deps.appState.config.imageGeneration = {
        enabled: true, provider: 'custom', providers: { custom: { ...config, model: 'image-fixture' } },
    };
    const detail = 'Second suggestion: a teal suspension bridge with three copper towers under a white sky.';
    deps.conversationHistory[1].content = 'Earlier discussion. '.repeat(100) + detail + 'Further discussion. '.repeat(100);
    const brief = 'A teal suspension bridge with three copper towers under a white sky.';
    fetch.mockResolvedValueOnce(response(brief))
        .mockResolvedValueOnce(fallback ? failure() : { ok: true, json: async () => ({ data: [{ b64_json: 'aW1hZ2U=' }] }) });
    if (fallback) fetch.mockResolvedValueOnce(response(svg));
    const instruction = 'Draw the second suggestion at document end';
    const result = await actions.prepareIllustrationProposal(deps, { instruction, onToken: jest.fn() });
    expect(result).toMatchObject({ renderer: fallback ? 'svg' : 'image', position: 'end', instruction });
    expect(fetch).toHaveBeenCalledTimes(fallback ? 3 : 2);
    const resolver = bodies()[0];
    expect(resolver.messages.map((m) => m.role)).toEqual(['system', 'user', 'assistant', 'user']);
    expect(resolver.messages.slice(1, -1)).toEqual(deps.conversationHistory);
    expect(resolver.messages[2].content).toContain(detail);
    expect(resolver.messages.at(-1).content).toContain(instruction);
    expect(resolver.stream).toBe(true);
    expect(fetch.mock.calls[0][0]).toContain('/chat/completions');
    expect(fetch.mock.calls[1][0]).toContain('/images/generations');
    expect(bodies()[1].prompt).toContain(brief);
    expect(bodies()[1].prompt).toContain('Current document');
    expect(bodies()[1].prompt).not.toContain('Further discussion');
    if (fallback) {
        expect(bodies()[2].messages.slice(0, -1)).toEqual(deps.conversationHistory);
        expect(bodies()[2].messages.at(-1).content).toContain(instruction);
    }
});

test('first-turn image generation makes no resolver call', async () => {
    const { deps } = world();
    deps.conversationHistory = [];
    deps.appState.config.imageGeneration = {
        enabled: true, provider: 'custom', providers: { custom: { ...config, model: 'image-fixture' } },
    };
    fetch.mockResolvedValueOnce({ ok: true, json: async () => ({ data: [{ b64_json: 'aW1hZ2U=' }] }) });
    const result = await actions.prepareIllustrationProposal(deps, { instruction: 'Draw a bridge at document end' });
    expect(result).toMatchObject({ renderer: 'image', position: 'end' });
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(fetch.mock.calls[0][0]).toContain('/images/generations');
    expect(bodies()[0].prompt).toContain('Draw a bridge');
});

test.each(['http', 'empty', 'unresolved', 'abort'])('image resolution failure stops generation and fallback: %s', async (kind) => {
    const { deps } = world();
    deps.appState.config.imageGeneration = {
        enabled: true, provider: 'custom', providers: { custom: { ...config, model: 'image-fixture' } },
    };
    const controller = new AbortController();
    if (kind === 'abort') {
        fetch.mockImplementationOnce(async () => { controller.abort(); return response('A bridge'); });
    } else {
        fetch.mockResolvedValueOnce(kind === 'http' ? failure() : response(kind === 'empty' ? '' : 'UNRESOLVED'));
    }
    const pending = actions.prepareIllustrationProposal(deps, {
        instruction: 'Draw that at document end', signal: controller.signal,
    });
    if (kind === 'abort') await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    else await expect(pending).rejects.toThrow(/Image request resolution failed/);
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(fetch.mock.calls[0][0]).toContain('/chat/completions');
    expect(deps.log.mock.calls.some(([message]) => message.includes('Falling back'))).toBe(false);
});

test.each([false, true])('every document chunk and retry retains one original history, stream=%s', async (stream) => {
    const { deps } = world();
    fetch.mockResolvedValueOnce(failure()).mockResolvedValueOnce(response('Revised two'));
    const run = await actions.runDocumentSkill(deps, {
        category: 'amendment', promptTemplate: 'Use the earlier approach', gateApply: true,
        ...(stream ? { onChunkToken: jest.fn() } : {}),
    });
    expect(run.failedCount).toBe(1);
    deps.conversationHistory = [{ role: 'user', content: 'A later unrelated turn' }];
    fetch.mockResolvedValueOnce(failure());
    const retry = await run.retryFailed();
    expect(retry.failedCount).toBe(1);
    fetch.mockResolvedValueOnce(response('Revised one'));
    const retried = await retry.retryFailed();
    expect(retried.failedCount).toBe(0);
    expect(fetch).toHaveBeenCalledTimes(4);
    for (const body of bodies()) expectHistory(body.messages, true);
});

test('comment capture and retry retain history after prompt validation', async () => {
    const { deps } = world();
    const queue = {
        addPending: jest.fn(() => 1), removePending: jest.fn(() => 0),
        captureSelectionAsBookmark: jest.fn(async () => {}),
        insertCommentOnBookmark: jest.fn(async () => ({ success: true, rangeText: 'Clause' })),
    };
    deps.appState.commentQueue = queue;
    fetch.mockResolvedValueOnce(failure()).mockResolvedValueOnce(response('Review comment'));
    await actions.fireSelectionComment(deps, { promptTemplate: 'Review that' });
    await settle();
    expect(deps.logWithRetry).toHaveBeenCalledTimes(1);
    deps.conversationHistory = [];
    deps.logWithRetry.mock.calls[0][2]();
    await settle();
    expect(queue.insertCommentOnBookmark).toHaveBeenCalledWith(expect.any(String), 'Review comment');
    expect(fetch).toHaveBeenCalledTimes(2);
    for (const body of bodies()) expectHistory(body.messages, true);
    expect(queue.captureSelectionAsBookmark).toHaveBeenCalledTimes(1);
});

test.each([[[]], [[system, ...history]]])('comment rejects invalid original message count before adding history: %j', (messages) => {
    const { deps } = world();
    resumeCommentFromBookmark('_one', '_one', 'Clause', {
        config, sendPromptFn: sendPrompt, conversationHistory: history,
        promptManager: { composeMessages: () => messages },
        commentQueue: { removePending: jest.fn(() => 0) },
        log: deps.log, updateStatusBarFn: jest.fn(),
    });
    expect(fetch).not.toHaveBeenCalled();
    expect(deps.log).toHaveBeenCalledWith('Comment failed: no comment prompt composed', 'error');
});

test('empty-history requests do not inherit a previous request history', async () => {
    const { deps } = world();
    await actions.answerQuestion(deps, { question: 'First', selectionText: 'Clause' });
    const fresh = world().deps;
    delete fresh.conversationHistory;
    await actions.answerQuestion(fresh, { question: 'New session', selectionText: 'Clause' });
    expect(bodies()[1].messages).toHaveLength(1);
    expect(JSON.stringify(bodies()[1])).not.toContain(history[1].content);
});

test('Anthropic Q&A keeps trusted system separate from prior assistant turns', async () => {
    const { deps } = world();
    deps.appState.config.backend = 'claude';
    deps.appState.config.providers.claude = { ...config, model: 'claude-sonnet-4-6' };
    fetch.mockResolvedValueOnce({
        ok: true, headers: { get: () => 'application/json' },
        json: async () => ({ content: [{ type: 'text', text: 'Continued answer' }], stop_reason: 'end_turn' }),
    });
    await expect(actions.answerQuestion(deps, { question: 'Continue that', selectionText: 'Clause' }))
        .resolves.toBe('Continued answer');
    const body = bodies()[0];
    expect(body.system).toBe(system.content);
    expectHistory(body.messages);
    expect(body.messages.at(-1).content).toContain('Continue that');
});

test('a document retry with originally absent history does not adopt later turns', async () => {
    const { deps } = world();
    delete deps.conversationHistory;
    chunkDocument.mockReturnValue([chunk('one')]);
    fetch.mockResolvedValueOnce(failure());
    const run = await actions.runDocumentSkill(deps, {
        category: 'amendment', promptTemplate: 'Revise', gateApply: true,
    });
    deps.conversationHistory = history;
    await run.retryFailed();
    expect(bodies()[1].messages.map((m) => m.role)).toEqual(['system', 'user']);
    expect(JSON.stringify(bodies()[1])).not.toContain(history[1].content);
});
