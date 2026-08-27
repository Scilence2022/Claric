/**
 * Specs for src/taskpane/agent-actions.js — the Word-side glue of the
 * tool-calling stack. All external layers are mocked: the LLM transport
 * (sendMessages drives the loop, sendPrompt the nested design call),
 * word-actions helpers, document extraction, and the Word global.
 */

jest.mock('../src/lib/llm-client.js', () => ({
    sendMessages: jest.fn(),
    sendPrompt: jest.fn(),
    sendPromptStream: jest.fn(),
    stripMarkdown: jest.fn((t) => t),
}));

jest.mock('../src/lib/illustration.js', () => ({
    buildIllustrationPrompt: jest.fn((instruction) => `design:${instruction}`),
    parseIllustration: jest.fn((raw) => ({ svg: String(raw) })),
    sanitizeSvg: jest.fn((svg) => svg),
    ensureSvgDimensions: jest.fn((svg) => svg),
    illustrationPositionFromInstruction: jest.fn(() => 'end'),
    illustrationPositionLabel: jest.fn(() => 'document end'),
    svgDimensions: jest.fn(() => ({ width: 1200, height: 800 })),
}));

jest.mock('../src/lib/comment-extractor.js', () => ({
    extractDocumentStructured: jest.fn(async () => 'DOCUMENT TEXT'),
}));

jest.mock('../src/taskpane/word-actions.js', () => ({
    readSelectionTableRegion: jest.fn(),
    svgToPngBase64: jest.fn(async () => ({ base64: 'cG5n', width: 1600, height: 900 })),
    insertPngPicture: jest.fn(() => ({ load: jest.fn(), width: 100, height: 50, altTextDescription: '' })),
    finalizeInsertedPicture: jest.fn(),
    imageDataUrl: jest.fn((b64) => String(b64).startsWith('data:')
        ? String(b64)
        : `data:image/png;base64,${b64}`),
}));

const { sendMessages, sendPrompt } = require('../src/lib/llm-client.js');
const { readSelectionTableRegion } = require('../src/taskpane/word-actions.js');
const { insertPngPicture, finalizeInsertedPicture } = require('../src/taskpane/word-actions.js');
const {
    prepareTableToolEdit, prepareImageToolEdit, applyImageOps,
} = require('../src/taskpane/agent-actions.js');

const REGION = {
    rowCount: 3,
    colCount: 2,
    values: [
        ['Header A', 'Header B'],
        ['old a', 'b'],
        ['c', 'd'],
    ],
    bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
    merged: false,
    shadowKeys: new Set(),
    cells: [
        { row: 1, col: 1, text: 'Header A' }, { row: 1, col: 2, text: 'Header B' },
        { row: 2, col: 1, text: 'old a' }, { row: 2, col: 2, text: 'b' },
        { row: 3, col: 1, text: 'c' }, { row: 3, col: 2, text: 'd' },
    ],
};

const SVG = '<svg width="1200" height="800" viewBox="0 0 1200 800"><rect width="1200" height="800"/></svg>';

function makeDeps() {
    return {
        appState: {
            config: {
                backend: 'mock',
                providers: { mock: { url: '', apiKey: '', model: 'mock-model', apiPath: '' } },
                trackChangesEnabled: true,
                docExtraction: { richness: 'structured' },
            },
        },
        log: jest.fn(),
    };
}

/** Word.run mock whose body lists the given picture proxies. */
function setImageWorld(pictures) {
    const items = pictures.map((over) => ({
        width: 300, height: 200, altTextDescription: '',
        load: jest.fn(), delete: jest.fn(),
        getRange: jest.fn(() => ({ insertInlinePictureFromBase64: jest.fn(() => ({ load: jest.fn() })) })),
        ...over,
    }));
    global.Word = {
        run: jest.fn(async (cb) => cb({
            document: { body: { inlinePictures: { items, load: jest.fn() } } },
            sync: jest.fn().mockResolvedValue(undefined),
        })),
        RangeLocation: { start: 'Start', end: 'End' },
        InsertLocation: { start: 'Start', end: 'End', before: 'Before' },
        ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
    };
    return items;
}

describe('prepareTableToolEdit', () => {
    beforeEach(() => jest.clearAllMocks());

    test('runs the loop against the region and returns a tablePatch proposal', async () => {
        readSelectionTableRegion.mockResolvedValue(REGION);
        sendMessages.mockResolvedValueOnce('{"tool":"set_cell","args":{"row":2,"col":1,"text":"new a"}}')
            .mockResolvedValueOnce('{"tool":"insert_row","args":{"position":"after","row":1,"values":["n1","n2"]}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"fixed and added a row"}}');

        const deps = makeDeps();
        const proposal = await prepareTableToolEdit(deps, { instruction: 'fix R2C1，然后加一行' });

        // Task prompt carries the grid; system prompt carries the tools.
        const messages = sendMessages.mock.calls[0][1];
        expect(messages[0].role).toBe('system');
        expect(messages[0].content).toContain('### set_cell');
        expect(messages[1].content).toContain('[R2C1] old a');
        expect(messages[1].content).toContain('fix R2C1，然后加一行');

        expect(proposal.tablePatch.cells).toEqual([{ row: 2, col: 1, text: 'new a' }]);
        expect(proposal.tablePatch.rowOps).toEqual([{ op: 'insertAfter', row: 1, values: ['n1', 'n2'] }]);
        expect(proposal.tableItems).toHaveLength(2);
        expect(proposal.tableItems[0]).toMatchObject({ label: 'Cell R2C1', before: 'old a', after: 'new a' });
        expect(proposal.toolLoop).toEqual({ steps: 3, finished: true });
    });

    test('returns null for non-table selections (caller falls back)', async () => {
        readSelectionTableRegion.mockResolvedValue(null);
        expect(await prepareTableToolEdit(makeDeps(), { instruction: 'x' })).toBeNull();
        expect(sendMessages).not.toHaveBeenCalled();
    });

    test('throws a noChanges error when the loop records nothing', async () => {
        readSelectionTableRegion.mockResolvedValue(REGION);
        sendMessages.mockResolvedValue('{"tool":"finish","args":{"summary":"nothing to do"}}');
        await expect(prepareTableToolEdit(makeDeps(), { instruction: 'x' }))
            .rejects.toMatchObject({ noChanges: true });
    });
});

describe('prepareImageToolEdit', () => {
    beforeEach(() => jest.clearAllMocks());

    test('drives design/delete/finish and returns ops + card items', async () => {
        setImageWorld([
            { altTextDescription: 'sunset' },
            {},
        ]);
        sendMessages.mockResolvedValueOnce('{"tool":"design_illustration","args":{"instruction":"a dusk scene","position":"end"}}')
            .mockResolvedValueOnce('{"tool":"delete_image","args":{"index":2}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"designed one, deleted one"}}');
        sendPrompt.mockResolvedValue(SVG);

        const deps = makeDeps();
        const proposal = await prepareImageToolEdit(deps, { instruction: '再加一张图，并删除第二张图片' });

        // Nested design call used the illustration prompt with doc context.
        expect(sendPrompt).toHaveBeenCalledWith(expect.anything(), 'design:a dusk scene', expect.anything(), undefined);

        expect(proposal.ops).toEqual([
            { type: 'insert', position: 'end', instruction: 'a dusk scene', svg: SVG },
            { type: 'delete', index: 2 },
        ]);
        expect(proposal.items).toHaveLength(2);
        expect(proposal.items[0]).toMatchObject({ id: 1, label: 'Insert illustration at end', svg: SVG });
        expect(proposal.items[1]).toMatchObject({ id: 2, label: expect.stringContaining('Delete image 2') });
        expect(proposal.snapshotCount).toBe(2);
    });

    test('throws a noChanges error when no ops and no usable summary', async () => {
        setImageWorld([]);
        sendMessages.mockResolvedValue('{"tool":"finish","args":{"summary":""}}');
        await expect(prepareImageToolEdit(makeDeps(), { instruction: 'x' }))
            .rejects.toMatchObject({ noChanges: true });
    });

    test('read-only loop (summary, no ops) resolves a noOps chat answer', async () => {
        setImageWorld([]);
        sendMessages.mockResolvedValue('{"tool":"finish","args":{"summary":"the image shows a bar chart"}}');
        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '这是什么图' });
        expect(proposal.noOps).toBe(true);
        expect(proposal.answer).toBe('the image shows a bar chart');
        expect(proposal.ops).toEqual([]);
        expect(proposal.items).toEqual([]);
    });

    test('read_image attaches the picture as a multimodal observation', async () => {
        setImageWorld([
            { altTextDescription: 'sunset', getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) },
        ]);
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"a sunset over water"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '描述这张图片' });

        // The send AFTER read_image carries an attachment-bearing observation.
// runToolLoop appends the latest assistant reply AFTER each send, so the
// snapshot is always one entry shorter than the array reference (which the
// mock captures) by the time the test inspects it.
        const messages = sendMessages.mock.calls[1][1];
        const observation = messages[messages.length - 2];
        expect(observation.role).toBe('user');
        expect(Array.isArray(observation.content)).toBe(true);
        expect(JSON.parse(observation.content[0].text)).toMatchObject({ ok: true, result: { index: 1 } });
        expect(observation.content[1]).toEqual({
            type: 'image_url',
            image_url: { url: 'data:image/png;base64,iVBORw0KGgo=' },
        });
        expect(proposal.noOps).toBe(true);
        expect(proposal.answer).toBe('a sunset over water');
    });

    test('read_image on a bad index becomes an error observation, not a crash', async () => {
        setImageWorld([{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]);
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":9}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"index missing"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '看第9张' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = messages[messages.length - 2];
        expect(observation.role).toBe('user');
        expect(typeof observation.content).toBe('string');
        expect(JSON.parse(observation.content).error).toMatch(/does not exist/);
        expect(proposal.answer).toBe('index missing');
    });

    test('selection metadata maps onto snapshot indexes in the task prompt', async () => {
        setImageWorld([
            { width: 300, height: 200, altTextDescription: 'sunset' },
            { width: 120, height: 80, altTextDescription: '' },
        ]);
        sendMessages.mockResolvedValue('{"tool":"finish","args":{"summary":""}}');

        await expect(prepareImageToolEdit(makeDeps(), {
            instruction: '删除这张图',
            selectionImages: [{ width: 300, height: 200, altText: 'sunset' }],
        })).rejects.toMatchObject({ noChanges: true });

        const messages = sendMessages.mock.calls[0][1];
        expect(messages[1].content).toContain('SELECTED by the user right now');
        expect(messages[1].content).toContain('image 1 (SELECTED');
        expect(messages[1].content).toContain('read_image');
    });

    test('unmatchable selection metadata degrades to an honest note', async () => {
        setImageWorld([{ width: 300, height: 200, altTextDescription: 'sunset' }]);
        sendMessages.mockResolvedValue('{"tool":"finish","args":{"summary":""}}');

        await expect(prepareImageToolEdit(makeDeps(), {
            instruction: 'x',
            selectionImages: [{ width: 1, height: 1, altText: 'nope' }],
        })).rejects.toMatchObject({ noChanges: true });

        const messages = sendMessages.mock.calls[0][1];
        expect(messages[1].content).toContain('could not be matched to a snapshot index');
    });

    test('HTTP 4xx with image parts retries once without attachments', async () => {
        setImageWorld([{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]);
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockRejectedValueOnce(new Error('HTTP 400: Bad Request'))
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"described from alt text"}}');

        const deps = makeDeps();
        const proposal = await prepareImageToolEdit(deps, { instruction: '看图' });

        expect(sendMessages).toHaveBeenCalledTimes(3);
        const retryMessages = sendMessages.mock.calls[2][1];
        // After the retry: the user-role observation is dropped to a text-only
        // array. The initial task user message has a string content and is not
        // in this filter, so finding any array-content user message proves
        // the strip pipeline ran.
        const strippedObservations = retryMessages.filter(
            (m) => m.role === 'user' && Array.isArray(m.content)
        );
        expect(strippedObservations.length).toBeGreaterThan(0);
        expect(strippedObservations[0].content.every((p) => p.type !== 'image_url')).toBe(true);
        expect(strippedObservations[0].content.length).toBeGreaterThan(0);
        expect(deps.log).toHaveBeenCalledWith(expect.stringMatching(/retrying without image attachments/), 'warning');
        expect(proposal.noOps).toBe(true);
        expect(proposal.answer).toBe('described from alt text');
    });

    test('abort during a loop send propagates without a retry', async () => {
        setImageWorld([{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]);
        const abortErr = new Error('The operation was aborted.');
        abortErr.name = 'AbortError';
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockRejectedValueOnce(abortErr);

        await expect(prepareImageToolEdit(makeDeps(), { instruction: '看图' }))
            .rejects.toMatchObject({ name: 'AbortError' });
        expect(sendMessages).toHaveBeenCalledTimes(2);
    });
});

describe('applyImageOps', () => {
    beforeEach(() => jest.clearAllMocks());

    test('index ops run first, inserts last; delete + insert verified', async () => {
        const items = setImageWorld([{ altTextDescription: 'sunset' }, {}]);
        const proposal = {
            snapshotCount: 2,
            ops: [
                { type: 'insert', position: 'cursor', instruction: 'a sun', svg: SVG },
                { type: 'delete', index: 2 },
                { type: 'altText', index: 1, text: 'chart' },
            ],
        };

        const result = await applyImageOps(makeDeps(), proposal);

        expect(items[1].delete).toHaveBeenCalled();
        expect(items[0].altTextDescription).toBe('chart');
        expect(items[0].delete).not.toHaveBeenCalled();
        expect(svgToPngBase64Mock()).toHaveBeenCalledWith(SVG);
        expect(insertPngPicture).toHaveBeenCalledWith(
            expect.anything(),
            expect.objectContaining({ base64: 'cG5n', position: 'cursor' })
        );
        expect(finalizeInsertedPicture).toHaveBeenCalledWith(
            expect.anything(), expect.anything(), 'a sun'
        );
        expect(result.applied).toBe(3);
        expect(result.warnings).toEqual([]);
    });

    test('aborts on snapshot-count drift (staleness guard)', async () => {
        setImageWorld([{}, {}]);
        const proposal = {
            snapshotCount: 3,
            ops: [{ type: 'delete', index: 1 }],
        };
        await expect(applyImageOps(makeDeps(), proposal))
            .rejects.toThrow(/changed since this proposal/);
    });

    test('empty ops reject', async () => {
        await expect(applyImageOps(makeDeps(), { ops: [] }))
            .rejects.toThrow(/No image operations/);
    });
});

/** svgToPngBase64 was imported via the word-actions mock. */
function svgToPngBase64Mock() {
    return require('../src/taskpane/word-actions.js').svgToPngBase64;
}
