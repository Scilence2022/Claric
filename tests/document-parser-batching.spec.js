/** @jest-environment jsdom */

/**
 * parseDocument reads revision-aware OOXML per paragraph. Each export is a
 * full package, so the reads must be spread over bounded syncs rather than
 * queued for the whole document at once.
 */
import { parseDocument, PARSE_BATCH_SIZE } from '../src/lib/document-parser.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const paragraphXml = (text) => `<w:document xmlns:w="${W}"><w:body>`
    + `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;

function mockDocument(texts, { failAtSync = null } = {}) {
    const state = { syncs: 0, pendingExports: 0, largestBatch: 0 };
    const items = texts.map((text, index) => ({
        text,
        style: 'Normal',
        styleBuiltIn: 'Normal',
        isListItem: false,
        load: jest.fn(),
        untrack: jest.fn(),
        parentTableOrNullObject: { isNullObject: index % 2 === 0, load: jest.fn() },
        listItemOrNullObject: { isNullObject: true, load: jest.fn() },
        getOoxml: jest.fn(() => {
            state.pendingExports++;
            return { value: paragraphXml(text) };
        }),
    }));
    const paragraphs = { items: [], load: jest.fn() };
    const context = {
        document: { body: { paragraphs } },
        sync: jest.fn(async () => {
            state.syncs++;
            if (state.syncs === failAtSync) {
                throw Object.assign(new Error('GeneralException'), { code: 'GeneralException' });
            }
            state.largestBatch = Math.max(state.largestBatch, state.pendingExports);
            state.pendingExports = 0;
            if (state.syncs === 1) paragraphs.items = items;
        }),
    };
    global.Word = { run: jest.fn(async (callback) => callback(context)) };
    return { state, items };
}

afterEach(() => {
    delete global.Word;
});

test('no sync carries more than one batch of paragraph OOXML exports', async () => {
    const texts = Array.from({ length: PARSE_BATCH_SIZE * 2 + 5 }, (_, i) => `Paragraph ${i + 1}`);
    const { state, items } = mockDocument(texts);

    const model = await parseDocument();

    expect(items.every((item) => item.getOoxml.mock.calls.length === 1)).toBe(true);
    expect(state.largestBatch).toBe(PARSE_BATCH_SIZE);
    // One enumeration sync plus three batches.
    expect(state.syncs).toBe(4);
    expect(model.paragraphs.map((p) => p.text)).toEqual(texts);
    expect(model.paragraphs.map((p) => p.inTable)).toEqual(texts.map((_, i) => i % 2 !== 0));
    expect(items.every((item) => item.untrack.mock.calls.length === 1)).toBe(true);
});

test('a failed batch names the paragraph range and keeps the Word error code', async () => {
    const texts = Array.from({ length: PARSE_BATCH_SIZE + 3 }, (_, i) => `Paragraph ${i + 1}`);
    mockDocument(texts, { failAtSync: 3 });

    await expect(parseDocument()).rejects.toMatchObject({
        message: expect.stringContaining(`paragraphs ${PARSE_BATCH_SIZE + 1}–${PARSE_BATCH_SIZE + 3}`),
        code: 'GeneralException',
    });
});
