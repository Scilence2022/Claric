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

function mockDocument(texts, { rejectBatch = () => false, reviewedFailure = -1 } = {}) {
    const state = { syncs: 0, largestBatch: 0, failedBatches: [] };
    const exports = [];
    const reviewed = [];
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
            exports.push(index);
            return { value: paragraphXml(text) };
        }),
        getRange: jest.fn(() => ({ getReviewedText: () => {
            reviewed.push(index);
            return { value: text };
        } })),
    }));
    const paragraphs = { items: [], load: jest.fn() };
    const context = {
        document: { body: { paragraphs } },
        sync: jest.fn(async () => {
            state.syncs++;
            const xmlBatch = exports.splice(0);
            const reviewedBatch = reviewed.splice(0);
            state.largestBatch = Math.max(state.largestBatch, xmlBatch.length);
            if (rejectBatch(xmlBatch) || reviewedBatch.includes(reviewedFailure)) {
                state.failedBatches.push(xmlBatch);
                throw Object.assign(new Error('GeneralException'), { code: 'GeneralException' });
            }
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

test('retries the reported paragraphs 97–120 in smaller reads and keeps document indexes', async () => {
    const texts = Array.from({ length: 125 }, (_, i) => `Paragraph ${i + 1}`);
    const { state, items } = mockDocument(texts, { rejectBatch: (batch) => batch[0] === 96 && batch.length === 24 });

    const model = await parseDocument();

    expect(state.failedBatches.map((batch) => [batch[0], batch.at(-1)])).toEqual([[96, 119]]);
    expect(model.paragraphs.map((p) => [p.index, p.text])).toEqual(texts.map((text, i) => [i, text]));
    expect(items[95].getOoxml).toHaveBeenCalledTimes(1);
    expect(items[96].getOoxml).toHaveBeenCalledTimes(2);
});

test('isolates an unavailable reviewed-text read and still uses OOXML revisions', async () => {
    const texts = ['Introduction.', 'Current text.', 'Conclusion.'];
    const { items } = mockDocument(texts, { reviewedFailure: 1 });
    items[1].text = 'Deleted text.Current text.';
    const originalExport = items[1].getOoxml.getMockImplementation();
    items[1].getOoxml.mockImplementation(() => ({ value: originalExport().value.replace('<w:p>',
        '<w:p><w:del><w:r><w:delText>Deleted text.</w:delText></w:r></w:del>') }));

    const model = await parseDocument();

    expect(model.paragraphs.map((p) => p.text)).toEqual(texts);
    expect(model.paragraphs[1]).toMatchObject({ hasRevisions: true, revisionFingerprint: expect.any(String) });
});

test('a persistent read failure names the exact paragraph and keeps the Word error code', async () => {
    const texts = Array.from({ length: PARSE_BATCH_SIZE + 3 }, (_, i) => `Paragraph ${i + 1}`);
    mockDocument(texts, { rejectBatch: (batch) => batch.includes(PARSE_BATCH_SIZE + 1) });

    await expect(parseDocument()).rejects.toMatchObject({
        message: expect.stringContaining(`paragraph ${PARSE_BATCH_SIZE + 2}`),
        code: 'GeneralException',
    });
});

test('refuses an unscoped XML export when reviewed text is unavailable', async () => {
    const { items } = mockDocument(['First paragraph.'], { reviewedFailure: 0 });
    const originalExport = items[0].getOoxml.getMockImplementation();
    items[0].getOoxml.mockImplementation(() => ({ value: originalExport().value.replace('</w:body>',
        '<w:p><w:r><w:t>Another paragraph.</w:t></w:r></w:p></w:body>') }));

    await expect(parseDocument()).rejects.toThrow(/paragraph 1: Word did not return XML scoped to one paragraph/);
});

test('cancellation during a failed sync stops before retrying any paragraph', async () => {
    const controller = new AbortController();
    const { items } = mockDocument(['One.', 'Two.'], { rejectBatch: (batch) => {
        if (!batch.length) return false;
        controller.abort();
        return true;
    } });

    await expect(parseDocument({ signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(items.every((item) => item.getOoxml.mock.calls.length === 1)).toBe(true);
});
