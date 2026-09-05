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

// Provide DOMParser for the OOXML text extraction (node test environment
// lacks it; the add-in WebView has it natively)
const { JSDOM } = require('jsdom');
if (typeof globalThis.DOMParser === 'undefined') {
    const dom = new JSDOM('');
    globalThis.DOMParser = dom.window.DOMParser;
}

jest.mock('../src/lib/illustration.js', () => ({
    buildIllustrationPrompt: jest.fn((instruction) => `design:${instruction}`),
    buildIllustrationRedesignPrompt: jest.fn((instruction, _scope, opts) => `redesign:${instruction}:${opts && opts.hasSourceImage ? 'with-image' : 'text-only'}${opts && opts.sourceSvg ? ':with-source' : ''}`),
    parseIllustration: jest.fn((raw) => ({ svg: String(raw) })),
    sanitizeSvg: jest.fn((svg) => svg),
    ensureSvgDimensions: jest.fn((svg) => svg),
    editSvgTextLabels: jest.fn((svg) => ({ svg, applied: [], failed: [], labels: [] })),
    illustrationPositionFromInstruction: jest.fn(() => 'end'),
    illustrationPositionLabel: jest.fn(() => 'document end'),
    svgDimensions: jest.fn(() => ({ width: 1200, height: 800 })),
}));

jest.mock('../src/taskpane/svg-source-store.js', () => ({
    SVG_SOURCE_TITLE_PREFIX: 'claric-svg:',
    attachSvgSource: jest.fn(async () => true),
    deleteSvgSource: jest.fn(async () => undefined),
    loadSvgSource: jest.fn(async () => null),
    svgSourceIdFromPicture: jest.fn((picture) => {
        const title = picture && picture.altTextTitle;
        return (typeof title === 'string' && title.startsWith('claric-svg:'))
            ? title.slice('claric-svg:'.length) : null;
    }),
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
const { buildIllustrationRedesignPrompt, editSvgTextLabels } = require('../src/lib/illustration.js');
const { attachSvgSource, deleteSvgSource, loadSvgSource } = require('../src/taskpane/svg-source-store.js');
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

/** Word.run mock whose body lists the given picture and paragraph proxies. */
function setImageWorld(pictures, { paragraphs, failContextSync = false } = {}) {
    const nullParagraph = {
        isNullObject: true,
        text: '',
        style: '',
        styleBuiltIn: '',
        load: jest.fn(),
    };
    nullParagraph.getPreviousOrNullObject = jest.fn(() => nullParagraph);
    nullParagraph.getNextOrNullObject = jest.fn(() => nullParagraph);

    const paragraphSpecs = paragraphs || pictures.map(() => ({}));
    const paragraphItems = paragraphSpecs.map((spec) => {
        const xmlText = String(spec.text || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
        const range = {
            getOoxml: jest.fn(() => ({
                value: `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>${xmlText}</w:t></w:r></w:p>`,
            })),
            insertText: jest.fn(),
        };
        return {
            isNullObject: false,
            text: '',
            style: 'Normal',
            styleBuiltIn: 'Normal',
            load: jest.fn(),
            getRange: jest.fn(() => range),
            ...spec,
        };
    });
    paragraphItems.forEach((paragraph, i) => {
        paragraph.getPreviousOrNullObject = jest.fn(() => paragraphItems[i - 1] || nullParagraph);
        paragraph.getNextOrNullObject = jest.fn(() => paragraphItems[i + 1] || nullParagraph);
        Object.defineProperty(paragraph, 'alignment', {
            get: () => paragraph._alignment,
            set: (v) => { paragraph._alignment = v; },
            configurable: true,
        });
        paragraph._alignment = paragraph.alignment || 'Left';
    });

    const items = pictures.map((over, i) => {
        const paragraph = paragraphItems[over.paragraphIndex === undefined ? i : over.paragraphIndex];
        if (paragraph && over.alignment) paragraph._alignment = over.alignment;
        return {
            width: 300, height: 200, altTextDescription: '', altTextTitle: '',
            hyperlink: '', lockAspectRatio: true,
            paragraph,
            load: jest.fn(), delete: jest.fn(),
            getRange: jest.fn(() => ({ insertInlinePictureFromBase64: jest.fn(() => ({ load: jest.fn() })) })),
            ...over,
        };
    });
    let runCount = 0;
    global.Word = {
        run: jest.fn(async (cb) => {
            runCount += 1;
            return cb({
                document: { load: jest.fn(), changeTrackingMode: 'TrackMineOnly', body: { inlinePictures: { items, load: jest.fn() } } },
                sync: failContextSync && runCount === 3
                    ? jest.fn().mockRejectedValue(new Error('context sync failed'))
                    : jest.fn().mockResolvedValue(undefined),
            });
        }),
        RangeLocation: { start: 'Start', end: 'End', content: 'Content' },
        InsertLocation: { start: 'Start', end: 'End', before: 'Before', replace: 'Replace' },
        ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
    };
    return items;
}

describe('prepareTableToolEdit', () => {
    beforeEach(() => jest.clearAllMocks());

    test('passes prior conversation to initial and subsequent table requests', async () => {
        const deps = makeDeps();
        deps.conversationHistory = [
            { role: 'user', content: 'What should the header say?' },
            { role: 'assistant', content: 'Use Revenue.' },
        ];
        const sent = [];
        const replies = ['{"tool":"get_state","args":{}}', '{"tool":"finish","args":{"summary":"reviewed"}}'];
        sendMessages.mockImplementation(async (_config, messages) => {
            sent.push(JSON.parse(JSON.stringify(messages)));
            return replies.shift();
        });
        await prepareTableToolEdit(deps, { instruction: 'check it again', region: REGION });
        expect(sent).toHaveLength(2);
        for (const messages of sent) {
            expect(messages[0].role).toBe('system');
            expect(messages.slice(1, 3)).toEqual(deps.conversationHistory);
            expect(messages[3].content).toContain('USER TASK: check it again');
        }
        expect(sent[1]).toHaveLength(6);
        sendMessages.mockReset();
    });

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

    test('merge_cells stages a merge op + card item', async () => {
        readSelectionTableRegion.mockResolvedValue(REGION);
        sendMessages.mockResolvedValueOnce('{"tool":"set_cell","args":{"row":2,"col":1,"text":"merged"}}')
            .mockResolvedValueOnce('{"tool":"merge_cells","args":{"row":2,"col":1,"rows":2,"cols":2}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"merged the block"}}');

        const proposal = await prepareTableToolEdit(makeDeps(), { instruction: '把 R2C1 到 R3C2 合并' });

        expect(proposal.tablePatch.merges).toEqual([
            { op: 'merge', startRow: 2, startCol: 1, endRow: 3, endCol: 2 },
        ]);
        // Card item for the merge, after the cell item.
        const mergeItem = proposal.tableItems.find((i) => /^Merge R2C1/.test(i.label));
        expect(mergeItem).toMatchObject({ before: 'old a | b | c | d', after: 'merged' });
    });

    test('merge-only result is a REAL proposal (noOps stays undefined)', async () => {
        readSelectionTableRegion.mockResolvedValue(REGION);
        sendMessages.mockResolvedValueOnce('{"tool":"merge_cells","args":{"row":1,"col":1,"rows":1,"cols":2}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"merged row 1"}}');

        const proposal = await prepareTableToolEdit(makeDeps(), { instruction: '合并第一行' });

        expect(proposal.noOps).toBeUndefined();
        expect(proposal.tablePatch.merges).toHaveLength(1);
        expect(proposal.tableItems).toHaveLength(1);
    });

    test('a read-only loop (get_state then finish) still resolves noOps', async () => {
        readSelectionTableRegion.mockResolvedValue(REGION);
        sendMessages.mockResolvedValueOnce('{"tool":"get_state","args":{}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"the table looks fine"}}');

        const proposal = await prepareTableToolEdit(makeDeps(), { instruction: '看看这个表格' });

        expect(proposal.noOps).toBe(true);
        expect(proposal.answer).toBe('the table looks fine');
    });

    test('returns null for non-table selections (caller falls back)', async () => {
        readSelectionTableRegion.mockResolvedValue(null);
        expect(await prepareTableToolEdit(makeDeps(), { instruction: 'x' })).toBeNull();
        expect(sendMessages).not.toHaveBeenCalled();
    });

    test('accepts an explicit region (document-scope table_management path)', async () => {
        // Caller passed the document's first-table region directly — the
        // selection reader is bypassed entirely and never invoked.
        readSelectionTableRegion.mockResolvedValue(null);
        sendMessages.mockResolvedValueOnce('{"tool":"set_table_style","args":{"style":"TableGrid"}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"styled"}}');
        const documentRegion = {
            ...REGION,
            bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
            style: { styleBuiltIn: 'TableGrid' },
        };

        const proposal = await prepareTableToolEdit(makeDeps(), { instruction: '样式', region: documentRegion });

        expect(readSelectionTableRegion).not.toHaveBeenCalled();
        expect(proposal.noOps).toBeUndefined();
        expect(proposal.tablePatch.cells).toEqual([]);
        expect(proposal.tablePatch.styleOps).toHaveLength(1);
        expect(proposal.tableItems[0].label).toBe('Table look: style → TableGrid');
        // The task prompt surfaces the ORIGINAL style snapshot from the region.
        const taskPrompt = sendMessages.mock.calls[0][1][1].content;
        expect(taskPrompt).toContain('Current table style:');
    });

    test('throws a noChanges error when the loop records nothing and produces no summary', async () => {
        readSelectionTableRegion.mockResolvedValue(REGION);
        sendMessages.mockResolvedValue('{"tool":"finish","args":{"summary":""}}');
        await expect(prepareTableToolEdit(makeDeps(), { instruction: 'x' }))
            .rejects.toMatchObject({ noChanges: true });
    });

    test('read-only loop (summary, no patch) resolves a noOps chat answer', async () => {
        readSelectionTableRegion.mockResolvedValue(REGION);
        sendMessages.mockResolvedValue('{"tool":"finish","args":{"summary":"row totals inconsistent"}}');
        const proposal = await prepareTableToolEdit(makeDeps(), { instruction: '看看这个表格' });
        expect(proposal.noOps).toBe(true);
        expect(proposal.answer).toBe('row totals inconsistent');
        expect(proposal.tablePatch).toBeUndefined();
        expect(proposal.tableItems).toBeUndefined();
    });

    test('style-only loop stages a real proposal with label-only card items', async () => {
        readSelectionTableRegion.mockResolvedValue({
            ...REGION,
            style: { styleBuiltIn: 'TableGrid', headerRowCount: 0, font: { bold: false } },
        });
        sendMessages.mockResolvedValueOnce('{"tool":"set_borders","args":{"borders":{"top":{"type":"single","width":1.5},"bottom":{"type":"single","width":1.5},"inside":"none"}}}')
            .mockResolvedValueOnce('{"tool":"set_header_row","args":{"rows":1,"font":{"bold":true}}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"three-line table applied"}}');

        const deps = makeDeps();
        const proposal = await prepareTableToolEdit(deps, { instruction: '改成三线表，表头加粗' });

        expect(proposal.noOps).toBeUndefined();
        expect(proposal.tablePatch.cells).toEqual([]);
        expect(proposal.tablePatch.styleOps).toHaveLength(2);
        // Card items are label-only descriptions (no before/after diff).
        expect(proposal.tableItems).toHaveLength(2);
        expect(proposal.tableItems[0].label).toMatch(/^Borders: top single 1\.5pt/);
        expect(proposal.tableItems[0]).not.toHaveProperty('before');
        expect(proposal.tableItems[1].label).toBe('Header: 1 header row(s), bold');
        // The task prompt carries the style snapshot and tool guidance.
        const taskPrompt = sendMessages.mock.calls[0][1][1].content;
        expect(taskPrompt).toContain('Current table style:');
        expect(taskPrompt).toContain('TableGrid');
        expect(taskPrompt).toContain('set_table_style');
    });
});

describe('prepareImageToolEdit', () => {
    beforeEach(() => jest.clearAllMocks());

    test.each(['design', 'redesign', 'redesign-fallback', 'redesign-unreadable'])(
        'nested %s requests explicitly include prior user and assistant context', async (kind) => {
            setImageWorld(kind === 'design' ? [] : [{
                getBase64ImageSrc: () => ({ value: kind === 'redesign-unreadable' ? '' : 'iVBORw0KGgo=' }),
            }]);
            const deps = makeDeps();
            deps.conversationHistory = [
                { role: 'user', content: 'Use a blue background.' },
                { role: 'assistant', content: 'I suggest a circular layout.' },
            ];
            const sent = [];
            const replies = [
                JSON.stringify({ tool: kind === 'design' ? 'design_illustration' : 'replace_illustration',
                    args: { index: 1, position: 'end', instruction: 'use that layout' } }),
                ...(kind === 'redesign-fallback' ? [new Error('HTTP 400: no vision')] : []),
                SVG,
                '{"tool":"finish","args":{"summary":"done"}}',
            ];
            sendMessages.mockImplementation(async (_config, messages) => {
                sent.push(JSON.parse(JSON.stringify(messages)));
                const reply = replies.shift();
                if (reply instanceof Error) throw reply;
                return reply;
            });
            const proposal = await prepareImageToolEdit(deps, { instruction: 'use that layout' });
            expect(proposal.ops).toHaveLength(1);
            expect(sent[0].slice(1, 3)).toEqual(deps.conversationHistory);
            expect(sent[sent.length - 1].slice(1, 3)).toEqual(deps.conversationHistory);
            for (const nested of sent.slice(1, -1)) {
                expect(nested.slice(0, 2)).toEqual(deps.conversationHistory);
                expect(nested).toHaveLength(3);
            }
            if (kind === 'redesign-fallback') {
                expect(sent[1][2].content[0].text).toBe('redesign:use that layout:with-image');
                expect(sent[2][2].content).toBe('redesign:use that layout:text-only');
            }
            expect(sendPrompt).not.toHaveBeenCalled();
            sendMessages.mockReset();
        }
    );

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

    test('read_image returns a following Caption paragraph as a strong candidate', async () => {
        setImageWorld(
            [{ paragraphIndex: 2, getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }],
            {
                paragraphs: [
                    { text: 'Earlier context' },
                    { text: 'Immediate context' },
                    { text: '   ' },
                    { text: '  Figure 2.   Revenue by region  ', style: 'Caption', styleBuiltIn: 'Caption' },
                    { text: 'Later discussion' },
                ],
            }
        );
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"caption read"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '读取图注' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content[0].text);
        expect(observation.result.documentContext.captionCandidates).toEqual([
            expect.objectContaining({
                text: 'Figure 2. Revenue by region',
                position: 'after',
                distance: 1,
                style: 'Caption',
                styleBuiltIn: 'Caption',
                captionStrength: 'strong',
                reason: 'Word built-in Caption style',
            }),
        ]);
        expect(observation.result.documentContext.nearbyParagraphs)
            .not.toEqual(expect.arrayContaining([expect.objectContaining({ position: 'containing' })]));
    });

    test('read_image recognizes a preceding Figure label as a weak candidate', async () => {
        setImageWorld(
            [{ paragraphIndex: 1, getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }],
            { paragraphs: [{ text: 'FIG. IV  Study flow' }, { text: '' }, { text: 'Body text' }] }
        );
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"candidate read"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '读取图片上下文' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content[0].text);
        expect(observation.result.documentContext.captionCandidates).toEqual([
            expect.objectContaining({
                text: 'FIG. IV Study flow',
                position: 'before',
                distance: 1,
                captionStrength: 'weak',
                reason: 'starts with a Figure/Fig./图 label and number',
            }),
        ]);
    });

    test('read_image then edit_figure_caption stages a visible Word-caption proposal', async () => {
        setImageWorld(
            [{ paragraphIndex: 0, getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }],
            { paragraphs: [
                { text: 'Figure image' },
                { text: 'Figure 3. Old legend', style: 'Caption', styleBuiltIn: 'Caption' },
                { text: 'Following text' },
            ] }
        );
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"edit_figure_caption","args":{'
                + '"index":1,"position":"after","distance":1,'
                + '"before":"Figure 3. Old legend","after":"Figure 3. Improved legend",'
                + '"evidence":{"strength":"strong","reason":"Word built-in Caption style"},'
                + '"style":{"style":"Caption","styleBuiltIn":"Caption"}}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"caption improved"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '评估 Figure legend，如有问题请改进' });

        expect(proposal.ops).toEqual([expect.objectContaining({
            type: 'figureCaption', index: 1, position: 'after', distance: 1,
            before: 'Figure 3. Old legend', after: 'Figure 3. Improved legend',
        })]);
        expect(proposal.items[0]).toMatchObject({
            before: 'Figure 3. Old legend', after: 'Figure 3. Improved legend',
        });
    });

    test('read_image handles null paragraph proxies at both document edges', async () => {
        setImageWorld(
            [{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }],
            { paragraphs: [{ text: 'Only paragraph', styleBuiltIn: 'Normal' }] }
        );
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"edge read"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '读取图片上下文' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content[0].text);
        expect(observation.result.documentContext).toMatchObject({
            nearbyParagraphs: [expect.objectContaining({
                text: 'Only paragraph', position: 'containing', distance: 0,
            })],
            captionCandidates: [],
            truncated: false,
        });
    });

    test('read_image normalizes and bounds paragraph context at word boundaries', async () => {
        const longText = (letter) => `${letter.repeat(600)}  \n\t ${letter.repeat(600)}`;
        setImageWorld(
            [{ paragraphIndex: 2, getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }],
            { paragraphs: ['a', 'b', 'c', 'd', 'e'].map((letter) => ({ text: longText(letter) })) }
        );
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"bounded"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '读取图片上下文' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content[0].text);
        const documentContext = observation.result.documentContext;
        expect(documentContext.truncated).toBe(true);
        // Normalized to 'aaaa…a bbbb…b' (600 + 1 + 600), capped at the last
        // space before 1000 → the 600-letter head plus an ellipsis.
        expect(documentContext.nearbyParagraphs).toHaveLength(5);
        for (const p of documentContext.nearbyParagraphs) {
            expect(p.truncated).toBe(true);
            expect(p.text).toMatch(/^[a-e]{600} …$/);
            expect(p.text.length).toBeLessThanOrEqual(1002);
            expect(p.text).not.toMatch(/[\n\t]| {2}/);
        }
    });

    test.each([
        ['unsupported nearby paragraph API', { paragraph: { load: jest.fn() } }, {}, /unavailable/],
        ['context sync failure', {}, { failContextSync: true }, /context sync failed/],
    ])('read_image keeps image data after %s', async (_label, pictureOverrides, options, reasonPattern) => {
        setImageWorld(
            [{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }), ...pictureOverrides }],
            options
        );
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"visual still available"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '看图' });

        const messages = sendMessages.mock.calls[1][1];
        const content = messages[messages.length - 2].content;
        const observation = JSON.parse(content[0].text);
        expect(observation.result.documentContext).toMatchObject({
            nearbyParagraphs: [], captionCandidates: [],
            unavailableReason: expect.stringMatching(reasonPattern),
        });
        expect(content[1]).toEqual({
            type: 'image_url',
            image_url: { url: 'data:image/png;base64,iVBORw0KGgo=' },
        });
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
        expect(proposal.answer).toContain('index missing');
        expect(proposal.answer).toContain('Unable to assess image pixels');
        expect(proposal.assessmentStatus).toBe('unable_to_assess');
    });

    test('read_image rejects unsupported bytes instead of claiming visual input is available', async () => {
        setImageWorld([{ getBase64ImageSrc: () => ({ value: 'Qk0AAAAA' }) }]); // BMP magic
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"unsupported image"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '检查图例' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content);
        expect(observation.error).toMatch(/format is not supported/);
        expect(proposal.assessmentStatus).toBe('unable_to_assess');
        expect(proposal.visualInputAvailable).toBe(false);
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
        expect(messages[1].content).toContain('legend or caption');
        expect(messages[1].content).toContain('untrusted data');
        expect(messages[1].content).toContain('do not claim an ordinary nearby paragraph');
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

    test('ambiguous selected-image metadata never defaults to the first picture', async () => {
        setImageWorld([
            { width: 300, height: 200, altTextDescription: 'same' },
            { width: 300, height: 200, altTextDescription: 'same' },
        ]);
        sendMessages
            .mockResolvedValueOnce('{"tool":"delete_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"selection is ambiguous"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), {
            instruction: '删除这张图',
            selectionImages: [{ width: 300, height: 200, altText: 'same' }],
        });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content);
        expect(observation.error).toMatch(/matches multiple document images/);
        expect(proposal.noOps).toBe(true);
        expect(proposal.ops).toEqual([]);
        expect(sendMessages.mock.calls[0][1][1].content).toContain('selection is ambiguous');
    });

    test('HTTP 4xx with image parts retries once without attachments', async () => {
        setImageWorld([{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]);
        let rejectedMessages;
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockImplementationOnce(async (_config, messages) => {
                rejectedMessages = JSON.parse(JSON.stringify(messages));
                throw new Error('HTTP 400: Bad Request');
            })
            .mockResolvedValueOnce('{"tool":"list_images","args":{}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"described from alt text"}}');

        const deps = makeDeps();
        deps.conversationHistory = [
            { role: 'user', content: 'Focus on the green series.' },
            { role: 'assistant', content: 'The green series represents revenue.' },
        ];
        const proposal = await prepareImageToolEdit(deps, { instruction: '看图' });

        expect(sendMessages).toHaveBeenCalledTimes(4);
        const retryMessages = sendMessages.mock.calls[2][1];
        for (const [, messages] of sendMessages.mock.calls) {
            expect(messages[0].role).toBe('system');
            expect(messages.slice(1, 3)).toEqual(deps.conversationHistory);
            expect(messages[3].content).toContain('看图');
        }
        expect(retryMessages.slice(0, 5)).toEqual(rejectedMessages.slice(0, 5));
        const originalObservation = JSON.parse(rejectedMessages[5].content[0].text);
        const retryObservation = JSON.parse(retryMessages[5].content[0].text);
        expect(retryObservation.result.documentContext).toEqual(originalObservation.result.documentContext);
        expect(retryObservation.result.note).toBe(originalObservation.result.note);
        expect(retryObservation.result.index).toBe(originalObservation.result.index);
        const laterMessages = sendMessages.mock.calls[3][1];
        expect(laterMessages[5]).toEqual(retryMessages[5]);
        expect(laterMessages.some((m) => Array.isArray(m.content)
            && m.content.some((part) => part.type === 'image_url'))).toBe(false);
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
        expect(proposal.answer).toContain('described from alt text');
        expect(proposal.answer).toContain('Unable to assess image pixels');
        expect(proposal.assessmentStatus).toBe('unable_to_assess');
        expect(proposal.visualInputAvailable).toBe(false);
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

    test('read_image resolves tracked-changes context via OOXML (accept-all view)', async () => {
        // Word.js paragraph.text inlines tracked deletions next to their
        // insertions ("TheResults reassemblerare maps resultsmapped...");
        // the OOXML path must hide w:del and keep w:ins + plain runs.
        const revisedOoxml =
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
            + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="text/xml"><pkg:xmlData>'
            + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
            + '<w:p>'
            + '<w:del><w:r><w:delText xml:space="preserve">The </w:delText></w:r></w:del>'
            + '<w:ins><w:r><w:t xml:space="preserve">Results </w:t></w:r></w:ins>'
            + '<w:del><w:r><w:delText xml:space="preserve">reassembler </w:delText></w:r></w:del>'
            + '<w:ins><w:r><w:t xml:space="preserve">are </w:t></w:r></w:ins>'
            + '<w:del><w:r><w:delText xml:space="preserve">maps results </w:delText></w:r></w:del>'
            + '<w:ins><w:r><w:t xml:space="preserve">mapped </w:t></w:r></w:ins>'
            + '<w:r><w:t xml:space="preserve">back onto the document.</w:t></w:r>'
            + '</w:p>'
            + '</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';
        setImageWorld(
            [{ paragraphIndex: 1, getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }],
            {
                paragraphs: [
                    {
                        text: 'The Results reassembler are maps results mapped back onto the document.',
                        getRange: jest.fn(() => ({ getOoxml: () => ({ value: revisedOoxml }) })),
                    },
                    { text: ' ' },
                    { text: 'Figure 1. Architecture', style: 'Caption', styleBuiltIn: 'Caption' },
                ],
            }
        );
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"clean context"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '读取图片上下文' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content[0].text);
        const texts = observation.result.documentContext.nearbyParagraphs.map((p) => p.text);
        expect(texts).toContain('Results are mapped back onto the document.');
        expect(texts.join('\n')).not.toContain('reassembler');
    });

    test('conditional Figure review cannot edit before read_image succeeds', async () => {
        setImageWorld([{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]);
        sendMessages
            .mockResolvedValueOnce('{"tool":"replace_illustration","args":{"index":1,"instruction":"improve legend"}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"unable to assess without reading"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), {
            instruction: '选择图像的 Figure legends 是否有改进的空间？如果有请改进',
        });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content);
        expect(observation.error).toMatch(/Figure legend or caption/);
        expect(sendPrompt).not.toHaveBeenCalled();
        expect(proposal.noOps).toBe(true);
        expect(proposal.assessmentStatus).toBe('unable_to_assess');
    });

    test('replace_illustration grounds a Figure edit in an explicit read_image', async () => {
        setImageWorld([{ altTextDescription: 'fig1', getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]);
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"replace_illustration","args":{"index":1,"instruction":"improve legends"}}')
            .mockResolvedValueOnce(SVG)
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"redesigned"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '改进这张图的图例' });

        // The nested design call follows the loop's read_image and replace calls.
        const designMessages = sendMessages.mock.calls[2][1];
        expect(designMessages).toHaveLength(1);
        expect(designMessages[0].role).toBe('user');
        expect(designMessages[0].content).toEqual([
            { type: 'text', text: 'redesign:improve legends:with-image' },
            { type: 'image_url', image_url: { url: 'data:image/png;base64,iVBORw0KGgo=' } },
        ]);
        expect(sendPrompt).not.toHaveBeenCalled();
        expect(proposal.ops).toEqual([
            {
                type: 'replace', index: 1, instruction: 'improve legends', svg: SVG,
                beforeSrc: 'data:image/png;base64,iVBORw0KGgo=',
            },
        ]);
    });

    test('Figure replacement refuses a blind redesign when the image is unreadable', async () => {
        setImageWorld([{ altTextDescription: 'fig1' }]); // no getBase64ImageSrc
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"replace_illustration","args":{"index":1,"instruction":"improve legends"}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"unable to assess"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '改进这张图的图例' });

        expect(sendPrompt).not.toHaveBeenCalled();
        expect(proposal.noOps).toBe(true);
        expect(proposal.ops).toEqual([]);
        expect(proposal.assessmentStatus).toBe('unable_to_assess');
    });

    test('Figure replacement does not fall back when the nested backend rejects pixels', async () => {
        setImageWorld([{ altTextDescription: 'fig1', getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]);
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"replace_illustration","args":{"index":1,"instruction":"improve legends"}}')
            .mockRejectedValueOnce(new Error('HTTP 400: Bad Request'))
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"redesign unavailable"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '改进这张图的图例' });

        expect(sendPrompt).not.toHaveBeenCalled();
        expect(proposal.noOps).toBe(true);
        expect(proposal.ops).toEqual([]);
    });

    test('replace_illustration feeds the stored SVG source to the design call', async () => {
        const OLD_SVG = '<svg width="10" height="10"><text>old label</text></svg>';
        setImageWorld([{
            altTextTitle: 'claric-svg:part-1',
            getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }),
        }]);
        loadSvgSource.mockResolvedValue(OLD_SVG);
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"replace_illustration","args":{"index":1,"instruction":"improve legends"}}')
            .mockResolvedValueOnce(SVG)
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"redesigned"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '改进这张图的图例' });

        // The nested design prompt was built with the exact source markup,
        // alongside the attached raster (see the mock's ':with-source' tag).
        expect(buildIllustrationRedesignPrompt).toHaveBeenCalledWith(
            'improve legends', expect.anything(),
            expect.objectContaining({ hasSourceImage: true, sourceSvg: OLD_SVG })
        );
        const designMessages = sendMessages.mock.calls[2][1];
        expect(designMessages[0].content[0].text).toBe('redesign:improve legends:with-image:with-source');
        expect(proposal.ops[0]).toEqual({
            type: 'replace', index: 1, instruction: 'improve legends', svg: SVG,
            beforeSrc: 'data:image/png;base64,iVBORw0KGgo=',
        });
        // The snapshot surfaces the capability without leaking the marker.
        const taskPrompt = sendMessages.mock.calls[0][1][1].content;
        expect(taskPrompt).toContain('editable SVG source');
        expect(taskPrompt).not.toContain('claric-svg:');
    });

    test('edit_illustration_text applies deterministic label edits without a design call', async () => {
        const OLD_SVG = '<svg width="10" height="10"><text>Dispatch</text></svg>';
        const EDITED_SVG = '<svg width="10" height="10"><text>调度</text></svg>';
        setImageWorld([{
            altTextTitle: 'claric-svg:part-1',
            getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }),
        }]);
        loadSvgSource.mockResolvedValueOnce(OLD_SVG);
        editSvgTextLabels.mockReturnValueOnce({
            svg: EDITED_SVG,
            applied: [{ old: 'Dispatch', new: '调度', count: 1 }],
            failed: [],
            labels: ['Dispatch'],
        });
        sendMessages
            .mockResolvedValueOnce('{"tool":"edit_illustration_text","args":{"index":1,"edits":[{"old":"Dispatch","new":"调度"}]}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"labels edited"}}');

        const proposal = await prepareImageToolEdit(makeDeps(), { instruction: '把图里的 Dispatch 改成 调度' });

        expect(editSvgTextLabels).toHaveBeenCalledWith(OLD_SVG, [{ old: 'Dispatch', new: '调度' }]);
        // Loop turns only — the edit never enters the nested design call.
        expect(sendMessages).toHaveBeenCalledTimes(2);
        expect(sendPrompt).not.toHaveBeenCalled();
        expect(proposal.ops).toEqual([{
            type: 'replace', index: 1,
            instruction: 'edit labels: "Dispatch" → "调度"',
            svg: EDITED_SVG,
            beforeSrc: 'data:image/png;base64,iVBORw0KGgo=',
        }]);
    });

    test('edit_illustration_text without a stored source points at replace_illustration', async () => {
        setImageWorld([{ getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }) }]); // no marker title
        sendMessages
            .mockResolvedValueOnce('{"tool":"edit_illustration_text","args":{"index":1,"edits":[{"old":"a","new":"b"}]}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"cannot edit directly"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '改图里的字' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content);
        expect(observation.error).toMatch(/no stored SVG source/);
        expect(observation.error).toMatch(/replace_illustration/);
        expect(editSvgTextLabels).not.toHaveBeenCalled();
    });

    test('edit_illustration_text reports the available labels when no edit matches', async () => {
        const OLD_SVG = '<svg width="10" height="10"><text>Alpha</text></svg>';
        setImageWorld([{
            altTextTitle: 'claric-svg:part-1',
            getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }),
        }]);
        loadSvgSource.mockResolvedValue(OLD_SVG);
        editSvgTextLabels.mockReturnValueOnce({
            svg: OLD_SVG,
            applied: [],
            failed: [{ old: 'Dispach', new: 'Dispatch' }],
            labels: ['Alpha', 'Beta'],
        });
        sendMessages
            .mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"edit_illustration_text","args":{"index":1,"edits":[{"old":"Dispach","new":"Dispatch"}]}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"retrying with the right label"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '修图例错字' });

        const messages = sendMessages.mock.calls[2][1];
        const observation = JSON.parse(messages[messages.length - 2].content);
        expect(observation.error).toMatch(/No edit matched/);
        expect(observation.error).toContain('"Alpha"');
    });

    test('read_image reports whether the picture carries a stored SVG source', async () => {
        setImageWorld([{
            altTextTitle: 'claric-svg:part-1',
            getBase64ImageSrc: () => ({ value: 'iVBORw0KGgo=' }),
        }]);
        loadSvgSource.mockResolvedValueOnce('<svg width="10" height="10"/>');
        sendMessages.mockResolvedValueOnce('{"tool":"read_image","args":{"index":1}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"seen"}}');

        await prepareImageToolEdit(makeDeps(), { instruction: '看图' });

        const messages = sendMessages.mock.calls[1][1];
        const observation = JSON.parse(messages[messages.length - 2].content[0].text);
        expect(observation.result.hasStoredSvgSource).toBe(true);
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

    test('resize by height keeps the width (unlocked)', async () => {
        const [pic] = setImageWorld([{ lockAspectRatio: false, width: 300, height: 200 }]);
        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'resize', index: 1, heightPt: 400, lockAspectRatio: false }],
        });
        expect(pic.height).toBe(400);
        expect(pic.width).toBe(600);
    });

    test('resize by scale computes both axes from the current size', async () => {
        const [pic] = setImageWorld([{ width: 200, height: 100 }]);
        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'resize', index: 1, scalePct: 150 }],
        });
        expect(pic.width).toBe(300);
        expect(pic.height).toBe(150);
    });

    test('altText sets both title and description independently', async () => {
        const [pic] = setImageWorld([{ altTextDescription: 'old' }]);
        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'altText', index: 1, title: 'Revenue chart', text: 'Q3 numbers' }],
        });
        expect(pic.altTextTitle).toBe('Revenue chart');
        expect(pic.altTextDescription).toBe('Q3 numbers');
    });

    test('align_image sets the picture paragraph alignment (Word string value)', async () => {
        const [pic] = setImageWorld([{ alignment: 'Left' }]);
        const result = await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'align', index: 1, alignment: 'centered' }],
        });
        expect(pic.paragraph._alignment).toBe('Centered');
        expect(result.applied).toBe(1);
        expect(result.warnings).toEqual([]);
    });

    test('set_image_link assigns the URL or null to clear', async () => {
        const [pic] = setImageWorld([{ hyperlink: 'https://old.example' }]);
        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'link', index: 1, url: 'https://new.example/report' }],
        });
        expect(pic.hyperlink).toBe('https://new.example/report');

        const [pic2] = setImageWorld([{ hyperlink: 'https://existing' }]);
        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'link', index: 1, url: null }],
        });
        expect([null, ''].includes(pic2.hyperlink)).toBe(true);
    });

    test('a failing image op warns and later ops still apply', async () => {
        // First picture has no paragraph anchor — the align op pushes a warning.
        setImageWorld([{ paragraph: null }, {}]);
        const result = await applyImageOps(makeDeps(), {
            snapshotCount: 2,
            ops: [
                { type: 'align', index: 1, alignment: 'centered' },
                { type: 'delete', index: 2 },
            ],
        });
        expect(result.applied).toBe(2);
        expect(result.warnings.some((w) => /no paragraph anchor/.test(w))).toBe(true);
    });

    test('insert persists the SVG source beside the new picture', async () => {
        setImageWorld([]);
        await applyImageOps(makeDeps(), {
            snapshotCount: 0,
            ops: [{ type: 'insert', position: 'end', instruction: 'a sun', svg: SVG }],
        });
        const insertedPic = insertPngPicture.mock.results[0].value;
        expect(attachSvgSource).toHaveBeenCalledWith(insertedPic, SVG);
    });

    test('replace scales and re-labels the new picture, re-attaches the new source, deletes the old part', async () => {
        const newPic = { load: jest.fn(), width: 1600, height: 900 };
        const [pic] = setImageWorld([{
            altTextTitle: 'claric-svg:old-part',
            altTextDescription: 'fig1',
            getRange: jest.fn(() => ({ insertInlinePictureFromBase64: jest.fn(() => newPic) })),
        }]);

        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'replace', index: 1, instruction: 'improve legends', svg: SVG }],
        });

        expect(pic.delete).toHaveBeenCalled();
        // The original alt-text description rides over to the replacement.
        expect(finalizeInsertedPicture).toHaveBeenCalledWith(expect.anything(), newPic, 'fig1');
        expect(attachSvgSource).toHaveBeenCalledWith(newPic, SVG);
        expect(deleteSvgSource).toHaveBeenCalledWith('old-part');
    });

    test('replace without a prior source still attaches the new one', async () => {
        const newPic = { load: jest.fn(), width: 1600, height: 900 };
        const [pic] = setImageWorld([{
            altTextDescription: '',
            getRange: jest.fn(() => ({ insertInlinePictureFromBase64: jest.fn(() => newPic) })),
        }]);

        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'replace', index: 1, instruction: 'redraw', svg: SVG }],
        });

        expect(attachSvgSource).toHaveBeenCalledWith(newPic, SVG);
        expect(deleteSvgSource).not.toHaveBeenCalled();
        expect(pic.delete).toHaveBeenCalled();
    });

    test('figureCaption applies only to the adjacent validated Caption paragraph', async () => {
        const [pic] = setImageWorld(
            [{ paragraphIndex: 0 }],
            { paragraphs: [
                { text: 'Figure image' },
                { text: 'Figure 4. Old legend', style: 'Caption', styleBuiltIn: 'Caption' },
            ] }
        );
        const caption = pic.paragraph.getNextOrNullObject();

        const result = await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{
                type: 'figureCaption', index: 1, position: 'after', distance: 1,
                before: 'Figure 4. Old legend', after: 'Figure 4. Improved legend',
                evidence: { captionStrength: 'strong', reason: 'Word built-in Caption style' },
                style: { style: 'Caption', styleBuiltIn: 'Caption' },
            }],
        });

        const writeRange = caption.getRange.mock.results.at(-1).value;
        expect(writeRange.insertText).toHaveBeenCalledWith('Figure 4. Improved legend', 'Replace');
        expect(result).toEqual({ applied: 1, warnings: [], interrupted: false, partial: false });
    });

    test('figureCaption skips a stale caption instead of writing another paragraph', async () => {
        const [pic] = setImageWorld(
            [{ paragraphIndex: 0 }],
            { paragraphs: [
                { text: 'Figure image' },
                { text: 'Figure 4. Changed by user', style: 'Caption', styleBuiltIn: 'Caption' },
            ] }
        );
        const caption = pic.paragraph.getNextOrNullObject();

        const result = await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{
                type: 'figureCaption', index: 1, position: 'after', distance: 1,
                before: 'Figure 4. Old legend', after: 'Figure 4. Improved legend',
                evidence: { captionStrength: 'strong', reason: 'Word built-in Caption style' },
                style: { style: 'Caption', styleBuiltIn: 'Caption' },
            }],
        });

        expect(caption.getRange.mock.results.every(({ value }) => !value.insertText.mock.calls.length)).toBe(true);
        expect(result.applied).toBe(0);
        expect(result.warnings[0]).toMatch(/changed since this proposal/);
    });

    test('figureCaption skips when OOXML is unavailable and fields cannot be ruled out', async () => {
        const [pic] = setImageWorld(
            [{ paragraphIndex: 0 }],
            { paragraphs: [
                { text: 'Figure image' },
                { text: 'Figure 4. Old legend', style: 'Caption', styleBuiltIn: 'Caption' },
            ] }
        );
        const caption = pic.paragraph.getNextOrNullObject();
        const unsafeRange = { insertText: jest.fn() };
        caption.getRange.mockImplementation(() => unsafeRange);

        const result = await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{
                type: 'figureCaption', index: 1, position: 'after', distance: 1,
                before: 'Figure 4. Old legend', after: 'Figure 4. Improved legend',
                evidence: { captionStrength: 'strong', reason: 'Word built-in Caption style' },
                style: { style: 'Caption', styleBuiltIn: 'Caption' },
            }],
        });

        expect(unsafeRange.insertText).not.toHaveBeenCalled();
        expect(result.applied).toBe(0);
        expect(result.warnings[0]).toMatch(/cannot be verified as safe plain text/);
    });

    test('delete drops the stored SVG source part', async () => {
        const [pic] = setImageWorld([{ altTextTitle: 'claric-svg:old-part' }]);
        await applyImageOps(makeDeps(), {
            snapshotCount: 1,
            ops: [{ type: 'delete', index: 1 }],
        });
        expect(pic.delete).toHaveBeenCalled();
        expect(deleteSvgSource).toHaveBeenCalledWith('old-part');
    });
});

/** svgToPngBase64 was imported via the word-actions mock. */
function svgToPngBase64Mock() {
    return require('../src/taskpane/word-actions.js').svgToPngBase64;
}

describe('prepareTableToolEdit (multi-table)', () => {
    beforeEach(() => jest.clearAllMocks());

    const SECOND_REGION = {
        rowCount: 2,
        colCount: 1,
        values: [['x'], ['y']],
        bounds: { startRow: 1, endRow: 2, startCol: 1, endCol: 1 },
        merged: false,
        shadowKeys: new Set(),
    };

    test('drives one loop over every region; tasks prompt lists both tables', async () => {
        sendMessages.mockResolvedValueOnce('{"tool":"set_table_style","args":{"tableIndex":2,"style":"TableGrid"}}')
            .mockResolvedValueOnce('{"tool":"set_cell","args":{"tableIndex":1,"row":2,"col":1,"text":"new a"}}')
            .mockResolvedValueOnce('{"tool":"finish","args":{"summary":"both tables done"}}');

        const proposal = await prepareTableToolEdit(makeDeps(), {
            instruction: '把两个表格都美化',
            regions: [REGION, SECOND_REGION],
        });

        expect(proposal.noOps).toBeUndefined();
        expect(proposal.tablePatch.tableCount).toBe(2);
        expect(proposal.tablePatch.cells).toEqual([
            { tableIndex: 1, row: 2, col: 1, text: 'new a' },
        ]);
        expect(proposal.tablePatch.styleOps).toEqual([
            { type: 'tableStyle', tool: 'set_table_style', tableIndex: 2, style: 'TableGrid' },
        ]);
        // Card items carry the table prefix.
        expect(proposal.tableItems[0].label).toBe('T1: Cell R2C1');
        expect(proposal.tableItems[1].label).toMatch(/^T2: Table look/);
        // Task prompt explains the multi-table coordinate scheme.
        const taskPrompt = sendMessages.mock.calls[0][1][1].content;
        expect(taskPrompt).toContain('table 1 (3x2):');
        expect(taskPrompt).toContain('table 2 (2x1):');
        expect(taskPrompt).toContain('tableIndex');
    });
});

describe('image write safety', () => {
    beforeEach(() => jest.clearAllMocks());
    test('stops at the next operation boundary and refuses to replay an attempted operation', async () => {
        const controller = new AbortController();
        const [first, second] = setImageWorld([{ delete: jest.fn(() => controller.abort()) }, {}]);
        const proposal = { snapshotCount: 2, ops: [{ type: 'delete', index: 1 }, { type: 'delete', index: 2 }] };
        const result = await applyImageOps(makeDeps(), proposal, { signal: controller.signal });
        expect(result).toMatchObject({ applied: 1, partial: true, interrupted: true });
        expect(first.delete).toHaveBeenCalledTimes(1);
        expect(second.delete).not.toHaveBeenCalled();
        await expect(applyImageOps(makeDeps(), proposal)).rejects.toThrow(/already attempted/);
    });
    test('checks cancellation after rasterization before inserting', async () => {
        const controller = new AbortController();
        const { svgToPngBase64 } = require('../src/taskpane/word-actions.js');
        svgToPngBase64.mockImplementationOnce(async () => { controller.abort(); return { base64: 'png' }; });
        setImageWorld([]);
        const result = await applyImageOps(makeDeps(), { snapshotCount: 0, ops: [{ type: 'insert', svg: SVG }] }, { signal: controller.signal });
        expect(result).toMatchObject({ applied: 0, interrupted: true, partial: false });
        expect(insertPngPicture).not.toHaveBeenCalled();
    });
    test('restores original tracking mode after an uncertain write failure', async () => {
        const [pic] = setImageWorld([{ delete: jest.fn(() => { throw new Error('host failure'); }) }]);
        const doc = { load: jest.fn(), changeTrackingMode: 'TrackMineOnly', body: { inlinePictures: { items: [pic], load: jest.fn() } } };
        Word.run.mockImplementation(async (fn) => fn({ document: doc, sync: async () => {} }));
        const result = await applyImageOps(makeDeps(), { snapshotCount: 1, ops: [{ type: 'delete', index: 1 }] });
        expect(result.partial).toBe(true);
        expect(doc.changeTrackingMode).toBe('TrackMineOnly');
    });
});
