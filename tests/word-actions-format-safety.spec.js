/** @jest-environment jsdom */
jest.mock('../src/lib/llm-client.js', () => ({ sendPrompt: jest.fn(), sendPromptStream: jest.fn() }));
jest.mock('../src/lib/comment-extractor.js', () => ({ extractDocumentStructured: jest.fn(async () => 'Document') }));
const { sendPrompt } = require('../src/lib/llm-client.js');
const { prepareFormatProposal, applyFormatProposal } = require('../src/taskpane/word-actions.js');

function world() {
    const range = { text: 'Same text', isNullObject: false, font: {}, load: jest.fn(), insertBookmark: jest.fn(), insertParagraph: jest.fn(() => ({ font: {} })) };
    const other = { text: 'Same text', font: {} };
    const document = {
        changeTrackingMode: 'TrackMineOnly', load: jest.fn(),
        getSelection: jest.fn(() => range),
        getBookmarkRangeOrNullObject: jest.fn(() => range),
        deleteBookmark: jest.fn(), body: { getRange: () => range },
    };
    const context = { document, sync: jest.fn(async () => {}) };
    global.Word = { run: jest.fn(async (fn) => fn(context)), ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' }, InsertLocation: { start: 'Start', end: 'End' } };
    const deps = { appState: { config: { backend: 'mock', providers: { mock: { model: 'm' } }, trackChangesEnabled: true } }, log: jest.fn() };
    return { range, other, document, context, deps };
}

beforeEach(() => { jest.clearAllMocks(); sendPrompt.mockResolvedValue('[{"font":{"bold":true}}]'); });
afterEach(() => { delete global.Word; });

test('applies only the captured bookmark, never an equal-text current selection', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'bold', selectionText: 'Same text' });
    expect(w.range.insertBookmark).toHaveBeenCalledWith(proposal.anchor.bookmark);
    w.document.getSelection.mockReturnValue(w.other);
    const result = await applyFormatProposal(w.deps, proposal);
    expect(w.range.font.bold).toBe(true);
    expect(w.other.font.bold).toBeUndefined();
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    expect(result.appliedRanges).toBe(1);
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
});

test.each(['missing', 'changed'])('refuses a %s anchor without falling back to equal text', async (kind) => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'bold' });
    if (kind === 'missing') w.range.isNullObject = true;
    else w.range.text = 'edited';
    w.document.getSelection.mockReturnValue(w.other);
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/changed or disappeared/);
    expect(w.range.font).toEqual({});
    expect(w.other.font).toEqual({});
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('cancellation after one insertion reports partial and never replays the insertion', async () => {
    const w = world();
    sendPrompt.mockResolvedValue('[{"insert":{"text":"Title","position":"start"}},{"font":{"bold":true}}]');
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'title then bold' });
    const controller = new AbortController();
    w.range.insertParagraph.mockImplementation(() => { controller.abort(); return { font: {} }; });
    const result = await applyFormatProposal(w.deps, proposal, { signal: controller.signal });
    expect(result).toMatchObject({ interrupted: true, partial: true, insertedParagraphs: 1 });
    expect(w.range.font).toEqual({});
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
    expect(w.range.insertParagraph).toHaveBeenCalledTimes(1);
});

test('restores tracking mode after a Word sync failure with uncertain writes', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'bold' });
    w.context.sync.mockResolvedValueOnce(undefined).mockRejectedValueOnce(new Error('Word failed')).mockResolvedValue(undefined);
    const result = await applyFormatProposal(w.deps, proposal);
    expect(result.partial).toBe(true);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('pre-aborted prepare neither anchors nor calls the model', async () => {
    const w = world();
    const controller = new AbortController(); controller.abort();
    await expect(prepareFormatProposal(w.deps, { instruction: 'bold', signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(w.range.insertBookmark).not.toHaveBeenCalled();
    expect(sendPrompt).not.toHaveBeenCalled();
});

test('rejects same-text formatting baseline drift when OOXML is available', async () => {
    const w = world();
    let xml = '<w:r><w:t>Same text</w:t></w:r>';
    w.range.getOoxml = jest.fn(() => ({ value: xml }));
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'bold' });
    xml = '<w:r><w:rPr><w:i/></w:rPr><w:t>Same text</w:t></w:r>';
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/baseline changed/);
    expect(w.range.font).toEqual({});
});

test('single illustration cancellation after insertion restores mode and prevents replay', async () => {
    const { applyIllustrationProposal } = require('../src/taskpane/word-actions.js');
    const w = world(); const controller = new AbortController();
    const picture = { width: 100, height: 100, load: jest.fn() };
    const insert = jest.fn(() => { controller.abort(); return picture; });
    w.document.body.insertParagraph = jest.fn(() => ({ getRange: () => ({ insertInlinePictureFromBase64: insert }) }));
    Word.RangeLocation = { start: 'Start' }; Word.Alignment = { centered: 'Centered' };
    const proposal = { imageBase64: 'png', position: 'end' };
    const result = await applyIllustrationProposal(w.deps, proposal, { signal: controller.signal });
    expect(result).toMatchObject({ inserted: true, partial: true, interrupted: true });
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyIllustrationProposal(w.deps, proposal)).rejects.toThrow(/already attempted/);
    expect(insert).toHaveBeenCalledTimes(1);
});

test('cursor illustration captures before generation and inserts at the original anchored selection end', async () => {
    const { prepareIllustrationProposal, applyIllustrationProposal } = require('../src/taskpane/word-actions.js');
    const w = world();
    w.range.getOoxml = jest.fn(() => ({ value: '<baseline/>' }));
    const insert = jest.fn(() => ({ width: 100, height: 50, load: jest.fn() }));
    w.range.getRange = jest.fn(() => ({ insertInlinePictureFromBase64: insert }));
    Word.RangeLocation = { start: 'Start', end: 'End' };
    sendPrompt.mockImplementation(async () => {
        expect(w.range.insertBookmark).toHaveBeenCalledTimes(1);
        w.document.getSelection.mockReturnValue(w.other);
        return '<svg width="100" height="50" viewBox="0 0 100 50"><rect width="100" height="50"/></svg>';
    });
    const proposal = await prepareIllustrationProposal(w.deps, { instruction: 'draw SVG at cursor' });
    expect(proposal.position).toBe('cursor');
    expect(proposal.anchor.ooxml).toBe('<baseline/>');
    await applyIllustrationProposal(w.deps, { ...proposal, imageBase64: 'png', svg: null });
    expect(w.document.getSelection).toHaveBeenCalledTimes(1);
    expect(w.range.getRange).toHaveBeenCalledWith('End');
    expect(insert).toHaveBeenCalledWith('png', 'Start');
});

test.each(['missing', 'changed', 'unsupported'])('cursor illustration refuses %s target evidence', async (kind) => {
    const { applyIllustrationProposal } = require('../src/taskpane/word-actions.js');
    const w = world();
    w.range.getOoxml = kind === 'unsupported' ? undefined : jest.fn(() => ({ value: kind === 'changed' ? '<new/>' : '<baseline/>' }));
    w.range.isNullObject = kind === 'missing';
    await expect(applyIllustrationProposal(w.deps, { imageBase64: 'png', position: 'cursor', anchor: { bookmark: '_cursor', text: 'Same text', ooxml: '<baseline/>' } })).rejects.toThrow(/anchor|baseline/);
    expect(w.document.getSelection).not.toHaveBeenCalled();
});

test('cursor prepare fails closed without OOXML capability', async () => {
    const { prepareIllustrationProposal } = require('../src/taskpane/word-actions.js');
    const w = world();
    await expect(prepareIllustrationProposal(w.deps, { instruction: 'draw at cursor' })).rejects.toThrow(/cannot safely anchor/);
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(w.range.insertBookmark).not.toHaveBeenCalled();
});
