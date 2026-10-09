/** @jest-environment jsdom */
jest.mock('../src/lib/llm-client.js', () => ({ sendPrompt: jest.fn(), sendPromptStream: jest.fn() }));
const { sendPrompt } = require('../src/lib/llm-client.js');
const { prepareFormatProposal, applyFormatProposal, discardFormatProposal } = require('../src/taskpane/word-actions.js');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const documentXml = (content) => `<w:document xmlns:w="${W}" xmlns:a="${A}"><w:body>${content}</w:body></w:document>`;
const textParagraph = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Example title</w:t></w:r></w:p>';
const table = '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr>'
    + '<w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
const drawing = '<w:p><w:r><w:drawing><a:graphic><a:graphicData uri="example"/></a:graphic></w:drawing></w:r></w:p>';
const field = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + '<w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
const original = textParagraph + table + drawing + field + '<w:p/>';

function world() {
    const state = { text: 'Example title\rCell\r\u0007\r\r', xml: documentXml(original) };
    const writes = [];
    const ranges = [];
    const makeBodyRange = () => {
        const range = {
            text: state.text, load: jest.fn(), getOoxml: jest.fn(() => ({ value: state.xml })), font: {},
            insertBookmark: jest.fn(() => { throw new Error('Mac cannot bookmark this complete body'); }),
        };
        Object.defineProperty(range.font, 'bold', { set(value) { writes.push(['bold', value]); }, configurable: true });
        ranges.push(range);
        return range;
    };
    const selection = { text: 'Example title', load: jest.fn(), font: {}, insertBookmark: jest.fn() };
    const bookmark = { text: 'Example title\rCell\r', isNullObject: false, load: jest.fn(), font: {} };
    const document = {
        changeTrackingMode: 'TrackMineOnly', load: jest.fn(), body: { getRange: jest.fn(makeBodyRange) },
        getSelection: jest.fn(() => selection), getBookmarkRangeOrNullObject: jest.fn(() => bookmark), deleteBookmark: jest.fn(),
    };
    const context = { document, sync: jest.fn(async () => {}) };
    global.Word = { run: jest.fn(async (fn) => fn(context)), ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' } };
    const deps = { appState: { config: { backend: 'mock', providers: { mock: { model: 'm' } }, trackChangesEnabled: true } }, log: jest.fn() };
    return { state, writes, ranges, selection, bookmark, document, context, deps };
}

beforeEach(() => { jest.clearAllMocks(); sendPrompt.mockResolvedValue('[{"font":{"bold":false}}]'); });
afterEach(() => { delete global.Word; });

test('whole-document formatting handles native tables, objects and terminal markers without a body bookmark', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix bold throughout the document', scope: 'document', selectionText: 'Example title' });
    expect(proposal.anchor).toMatchObject({ kind: 'document-body', text: w.state.text, ooxml: w.state.xml });
    expect(proposal.anchor.structureFingerprint).toBeTruthy();
    expect(proposal.anchor.bookmark).toBeNull();
    expect(sendPrompt.mock.calls[0][1]).toContain('Cell');
    w.selection.text = 'Other selected paragraph';
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, appliedRanges: 1, partial: false });
    expect(w.writes).toEqual([['bold', false]]);
    expect(w.document.body.getRange.mock.calls).toEqual([['Whole'], ['Whole']]);
    for (const range of w.ranges) expect(range.insertBookmark).not.toHaveBeenCalled();
    expect(w.document.getSelection).not.toHaveBeenCalled();
    expect(w.document.getBookmarkRangeOrNullObject).not.toHaveBeenCalled();
    expect(w.document.deleteBookmark).not.toHaveBeenCalled();
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('document formatting works when native bookmark APIs are unavailable', async () => {
    const w = world();
    delete w.document.getBookmarkRangeOrNullObject;
    delete w.document.deleteBookmark;
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, appliedRanges: 1 });
});

test('document cleanup uses native body containment and preserves table, picture and final paragraphs', async () => {
    const w = world();
    const lastRange = {};
    const makeParagraph = (kind) => {
        const nativeRange = { compareLocationWith: jest.fn((other) => ({ value: other === lastRange && kind === 'final' ? 'Equal' : 'Inside' })) };
        return {
            text: '', getRange: () => nativeRange, delete: jest.fn(),
            parentTableOrNullObject: { isNullObject: kind !== 'table', load: jest.fn() },
            getOoxml: () => ({ value: `<w:p xmlns:w="${W}">${kind === 'picture' ? '<w:r><w:drawing/></w:r>' : ''}</w:p>` }),
        };
    };
    const paragraphs = ['blank', 'table', 'picture', 'final'].map(makeParagraph);
    const createRange = w.document.body.getRange.getMockImplementation();
    w.document.body.getRange.mockImplementation(() => ({ ...createRange(), paragraphs: { items: paragraphs, load: jest.fn() } }));
    w.document.body.paragraphs = { getLast: () => ({ getRange: () => lastRange }) };
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Remove redundant empty paragraphs throughout the document', scope: 'document', cleanupOnly: true });
    expect(proposal.cleanupSummary).toMatchObject({ candidates: 4, verified: 1, preserved: 3, unverifiable: 0 });
    expect(proposal.ops).toEqual([{ cleanup: { emptyParagraphs: true, emptyCount: 1 } }]);
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, deletedParagraphs: 1, partial: false });
    expect(paragraphs[0].delete).toHaveBeenCalledTimes(1);
    for (const paragraph of paragraphs.slice(1)) expect(paragraph.delete).not.toHaveBeenCalled();
    expect(w.document.getBookmarkRangeOrNullObject).not.toHaveBeenCalled();
});

test.each([
    ['paragraph style', original.replace('Heading1', 'Normal')],
    ['table borders', original.replace('w:val="single"', 'w:val="double"')],
    ['drawing', original.replace('uri="example"', 'uri="changed"')],
    ['field', original.replace(' PAGE ', ' NUMPAGES ')],
    ['tracked content', original.replace('<w:t>Cell</w:t>', '<w:delText>Cell</w:delText>')],
    ['extra paragraph', original + '<w:p/>'],
])('same-text %s changes invalidate document formatting before any write', async (_change, content) => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' });
    w.state.xml = documentXml(content);
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/baseline changed/);
    expect(w.writes).toEqual([]);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    expect(w.deps.log).toHaveBeenCalledWith(expect.stringContaining('baseline v3 mismatch (document)'), 'warning');
});

test('text changed during model planning rejects the stale whole-document proposal', async () => {
    const w = world();
    sendPrompt.mockImplementation(async () => {
        w.state.text += 'New paragraph';
        return '[{"font":{"bold":false}}]';
    });
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' });
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/changed or disappeared/);
    expect(w.writes).toEqual([]);
    expect(w.document.getSelection).not.toHaveBeenCalled();
});

test('Word rendering/session metadata changes preserve valid document formatting', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' });
    w.state.xml = documentXml(original.replace('<w:p>', '<w:p w:rsidR="12345678"><w:proofErr w:type="spellStart"/><w:lastRenderedPageBreak/>'));
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, appliedRanges: 1 });
});

test.each(['missing body', 'missing OOXML', 'invalid OOXML'])('%s fails closed before model planning', async (kind) => {
    const w = world();
    if (kind === 'missing body') w.document.body = undefined;
    if (kind === 'missing OOXML') w.document.body.getRange.mockReturnValue({ text: w.state.text, load: jest.fn() });
    if (kind === 'invalid OOXML') w.state.xml = '<invalid';
    await expect(prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' })).rejects.toThrow(/document formatting/);
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(w.writes).toEqual([]);
    expect(w.document.getSelection).not.toHaveBeenCalled();
    expect(w.document.deleteBookmark).not.toHaveBeenCalled();
});

test('document baseline read failure preserves the native error and creates no bookmark', async () => {
    const w = world();
    w.document.body.getRange.mockImplementation(() => ({ text: w.state.text, load: jest.fn(), getOoxml: () => { throw new Error('GeneralException'); } }));
    await expect(prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' })).rejects.toThrow('GeneralException');
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(w.document.deleteBookmark).not.toHaveBeenCalled();
});

test('missing body at apply never redirects a document proposal into the current selection', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' });
    w.document.body = undefined;
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/document formatting scope/);
    expect(w.document.getSelection).not.toHaveBeenCalled();
    expect(w.writes).toEqual([]);
});

test('discarded document proposal cannot be applied or replayed and needs no native cleanup', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' });
    await discardFormatProposal(w.deps, proposal);
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/discarded/);
    expect(w.writes).toEqual([]);
    expect(w.document.deleteBookmark).not.toHaveBeenCalled();
});

test('a document body anchor cannot be substituted for an authorized selection scope', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix document bold', scope: 'document' });
    proposal.scope = 'selection';
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/not anchored/);
    expect(w.writes).toEqual([]);
});

test('partial selection bookmark recovery still fails instead of expanding to the document', async () => {
    const w = world();
    w.selection.text = 'Example title\rCell';
    w.bookmark.text = 'Example title';
    await expect(prepareFormatProposal(w.deps, { instruction: 'Fix selected bold', scope: 'selection' })).rejects.toThrow(/exact formatting scope/);
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(w.document.body.getRange).not.toHaveBeenCalled();
    expect(w.document.deleteBookmark).toHaveBeenCalledTimes(1);
});
