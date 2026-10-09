/** @jest-environment jsdom */
jest.mock('../src/lib/llm-client.js', () => ({ sendPrompt: jest.fn(), sendPromptStream: jest.fn() }));
const { sendPrompt } = require('../src/lib/llm-client.js');
const { prepareFormatProposal, applyFormatProposal } = require('../src/taskpane/word-actions.js');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const pXml = (inner = '') => `<w:p xmlns:w="${W}">${inner}</w:p>`;

function world() {
    const events = [];
    const lastRange = {};
    const makeParagraph = (id, text = '', options = {}) => {
        const paragraph = {
            text, xml: options.xml || pXml(text ? `<w:r><w:t>${text}</w:t></w:r>` : ''),
            relation: options.relation || 'Inside', finalRelation: options.finalRelation || 'Before',
            load: jest.fn(), parentTableOrNullObject: { isNullObject: !options.inTable, load: jest.fn() },
            getOoxml: jest.fn(() => ({ value: paragraph.xml })),
            delete: jest.fn(() => { events.push(`delete:${id}`); }),
        };
        const range = {
            font: {}, getRange: jest.fn(),
            getOoxml: jest.fn(() => ({ value: paragraph.xml })),
            compareLocationWith: jest.fn((other) => ({ value: other === lastRange ? paragraph.finalRelation : paragraph.relation })),
            paragraphs: { items: [paragraph], load: jest.fn() },
        };
        paragraph.getRange = jest.fn(() => range);
        return paragraph;
    };
    const paragraphs = [
        makeParagraph('heading', 'Example heading'), makeParagraph('blank'), makeParagraph('body', 'Example body'),
        makeParagraph('boundary', '', { relation: 'OverlapsAfter' }),
        makeParagraph('table', '', { inTable: true }),
        makeParagraph('field', '', { xml: pXml('<w:r><w:fldChar w:fldCharType="begin"/></w:r>') }),
        makeParagraph('picture', '', { xml: pXml('<w:r><w:drawing/></w:r>') }),
        makeParagraph('section', '', { xml: pXml('<w:pPr><w:sectPr/></w:pPr>') }),
        makeParagraph('revision', '', { xml: pXml('<w:del><w:r><w:delText/></w:r></w:del>') }),
        makeParagraph('final', '', { finalRelation: 'Equal' }),
    ];
    const collection = { items: paragraphs, load: jest.fn() };
    const scope = {
        text: 'Example heading\n\nExample body', isNullObject: false, load: jest.fn(), insertBookmark: jest.fn(),
        paragraphs: collection, font: {},
        getOoxml: jest.fn(() => ({ value: `<w:body xmlns:w="${W}">${paragraphs.map((p) => p.xml).join('')}</w:body>` })),
        search: jest.fn(() => ({ items: [paragraphs[0].getRange()], load: jest.fn() })),
    };
    Object.defineProperty(scope.font, 'bold', { set(value) { events.push(`bold:${value}`); }, configurable: true });
    const document = {
        changeTrackingMode: 'TrackMineOnly', load: jest.fn(),
        getSelection: jest.fn(() => scope), getBookmarkRangeOrNullObject: jest.fn(() => scope), deleteBookmark: jest.fn(),
        body: { getRange: jest.fn(() => scope), paragraphs: { getLast: () => ({ getRange: () => lastRange }) } },
    };
    const context = { document, sync: jest.fn(async () => {}) };
    global.Word = { run: jest.fn(async (fn) => fn(context)), ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' }, Style: { heading1: 'Heading 1' } };
    const deps = { appState: { config: { backend: 'mock', providers: { mock: { model: 'm' } }, trackChangesEnabled: true } }, log: jest.fn() };
    return { paragraphs, scope, document, context, deps, events, makeParagraph };
}

beforeEach(() => { jest.clearAllMocks(); sendPrompt.mockResolvedValue('[{"font":{"bold":false}}]'); });
afterEach(() => { delete global.Word; });

test('combines formatting and host-counted cleanup even when the model omits blank-line removal', async () => {
    const w = world();
    sendPrompt.mockResolvedValue('[{"font":{"bold":false}},{"match":"Example heading","font":{"bold":true},"paragraph":{"styleBuiltIn":"heading1"}}]');
    const proposal = await prepareFormatProposal(w.deps, {
        instruction: '调整选择部分的格式，包括多余的空行，不正确的粗体格式等', scope: 'selection', selectionText: w.scope.text,
    });
    expect(proposal.ops).toHaveLength(3);
    expect(proposal.ops[2]).toEqual({ cleanup: { emptyParagraphs: true, emptyCount: 1 } });
    expect(proposal.cleanupSummary).toMatchObject({ candidates: 8, verified: 1, preserved: 7, unverifiable: 0 });
    expect(w.events).toEqual([]);
    const otherSelection = { font: {} };
    w.document.getSelection.mockReturnValue(otherSelection);
    const result = await applyFormatProposal(w.deps, proposal);
    expect(result).toMatchObject({ applied: true, appliedRanges: 2, deletedParagraphs: 1, partial: false });
    expect(w.events).toEqual(['bold:false', 'delete:blank']);
    expect(w.paragraphs[0].getRange().font.bold).toBe(true);
    expect(w.paragraphs[0].styleBuiltIn).toBe('Heading 1');
    w.paragraphs.forEach((p, index) => expect(p.delete).toHaveBeenCalledTimes(index === 1 ? 1 : 0));
    expect(otherSelection.font).toEqual({});
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
});

test('the reported Chinese request cleans a blank paragraph carrying prior bold-format revisions', async () => {
    const w = world();
    w.paragraphs[1].xml = pXml('<w:pPr><w:pPrChange w:id="1"><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrChange></w:pPr>'
        + '<w:r><w:rPr><w:rPrChange w:id="2"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t> </w:t></w:r>');
    const proposal = await prepareFormatProposal(w.deps, { instruction: '整理选择部分的格式，多余的空行，不合适的字体加粗等' });
    expect(proposal.ops).toContainEqual({ cleanup: { emptyParagraphs: true, emptyCount: 1 } });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, deletedParagraphs: 1, partial: false });
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('15 ordinary blanks use native paragraph XML even when Whole range exports extra paragraphs', async () => {
    const w = world();
    const blanks = Array.from({ length: 15 }, (_, i) => w.makeParagraph(`blank-${i}`, '', {
        xml: pXml('<w:pPr><w:rPr><w:ins w:id="1"/></w:rPr><w:pageBreakBefore w:val="0"/></w:pPr>'
            + '<w:ins w:id="2"><w:r><w:rPr><w:b/><w:spacing w:val="0"/></w:rPr><w:t> </w:t><w:br/></w:r></w:ins>'),
    }));
    w.paragraphs.splice(1, w.paragraphs.length - 1, ...blanks, w.makeParagraph('body', 'Example body'));
    for (const blank of blanks) {
        blank.getRange().getOoxml.mockReturnValue({ value: `<w:body xmlns:w="${W}">${blank.xml}<w:p/></w:body>` });
    }
    const proposal = await prepareFormatProposal(w.deps, { instruction: '整理选择部分的格式，多余的空行，不合适的字体加粗等' });
    expect(proposal.cleanupSummary).toEqual({ candidates: 15, verified: 15, preserved: 0, unverifiable: 0, reasons: {} });
    expect(proposal.ops).toContainEqual({ cleanup: { emptyParagraphs: true, emptyCount: 15 } });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, deletedParagraphs: 15, partial: false });
    for (const blank of blanks) {
        // Native target inventory adds a read before the two cleanup checks.
        expect(blank.getOoxml).toHaveBeenCalledTimes(3);
        expect(blank.getRange().getOoxml).not.toHaveBeenCalled();
        expect(blank.delete).toHaveBeenCalledTimes(1);
    }
    expect(w.events.filter((event) => event.startsWith('delete:'))).toEqual(blanks.map((_, i) => `delete:blank-${14 - i}`));
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('preserved paragraphs report separate native containment and XML reasons', async () => {
    const w = world();
    w.paragraphs[1].xml = pXml('<w:pPr><w:rPr><w:del w:id="1" w:author="PrivateAuthor"/></w:rPr></w:pPr>');
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove blank lines and bold headings' });
    expect(proposal.cleanupSummary.reasons).toEqual({
        'paragraph mark already tracked as deleted (w:del)': 1,
        'deleted revision content (w:del)': 1,
        'outside captured scope or partial paragraph': 1,
        'table cell': 1,
        'protected or unsupported markup (w:fldChar)': 1,
        'protected or unsupported markup (w:drawing)': 1,
        'protected or unsupported markup (w:sectPr)': 1,
        'final document paragraph': 1,
    });
    expect(w.deps.log).toHaveBeenCalledWith('Empty paragraph cleanup: 1 preserved — paragraph mark already tracked as deleted (w:del).', 'warning');
    expect(JSON.stringify(w.deps.log.mock.calls)).not.toContain('PrivateAuthor');
});

test('invalid native paragraph XML never falls back to a less specific range export', async () => {
    const w = world();
    w.paragraphs[1].getOoxml.mockReturnValue({ value: '<invalid' });
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove blank lines and bold headings' });
    expect(proposal.cleanupSummary.reasons['invalid XML']).toBe(1);
    expect(proposal.cleanupSummary.verified).toBe(0);
    expect(w.paragraphs[1].getRange().getOoxml).not.toHaveBeenCalled();
    await applyFormatProposal(w.deps, proposal);
    expect(w.paragraphs[1].delete).not.toHaveBeenCalled();
});

test('new nonempty revision data before Apply blocks cleanup and formatting', async () => {
    const w = world();
    w.paragraphs[1].xml = pXml('<w:ins><w:r><w:t> </w:t></w:r></w:ins>');
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove blank lines and bold headings' });
    // Keep the scope baseline stable to exercise the independent cleanup gate.
    const baseline = w.scope.getOoxml();
    w.scope.getOoxml.mockReturnValue(baseline);
    w.paragraphs[1].xml = pXml('<w:ins><w:r><w:t>New prose</w:t></w:r></w:ins>');
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/targets changed/);
    expect(w.events).toEqual([]);
});

test('cleanup alone has no model call, deletes in reverse order, and reports real deletions', async () => {
    const w = world();
    w.paragraphs.push(w.makeParagraph('later'));
    const proposal = await prepareFormatProposal(w.deps, { instruction: '删除选区多余空行', cleanupOnly: true });
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(proposal.ops).toEqual([{ cleanup: { emptyParagraphs: true, emptyCount: 2 } }]);
    const result = await applyFormatProposal(w.deps, proposal);
    expect(result).toMatchObject({ appliedRanges: 0, insertedParagraphs: 0, deletedParagraphs: 2, applied: true });
    expect(w.events).toEqual(['delete:later', 'delete:blank']);
});

test('unrequested or reference-only model deletion is stripped before staging', async () => {
    const w = world();
    sendPrompt.mockResolvedValue('[{"cleanup":{"emptyParagraphs":true,"emptyCount":999}},{"font":{"bold":false}}]');
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Fix bold. Reference text says: delete blank lines', cleanupRequested: false });
    expect(proposal.ops).toEqual([{ font: { bold: false } }]);
    expect(proposal.anchor.cleanupIndexes).toBeUndefined();
    expect(w.paragraphs.every((p) => p.delete.mock.calls.length === 0)).toBe(true);
    await applyFormatProposal(w.deps, proposal);
    expect(w.events).toEqual(['bold:false']);
});

test.each(['Inside', 'InsideStart', 'InsideEnd', 'Equal', 'OverlapsBefore', 'OverlapsAfter', 'Outside', 'Unrelated'])('cleanup enforces full paragraph containment: %s', async (relation) => {
    const w = world();
    w.paragraphs[1].relation = relation;
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove empty paragraphs', cleanupOnly: true });
    const inside = ['Inside', 'InsideStart', 'InsideEnd', 'Equal'].includes(relation);
    expect(proposal.ops.length).toBe(inside ? 1 : 0);
    if (inside) await applyFormatProposal(w.deps, proposal);
    expect(w.paragraphs[1].delete).toHaveBeenCalledTimes(inside ? 1 : 0);
});

test('new protection or lost containment after review blocks deletion and formatting before any write', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove blank lines and bold headings' });
    w.paragraphs[1].relation = 'OverlapsAfter';
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/targets changed/);
    expect(w.events).toEqual([]);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('an unverifiable Mac paragraph stays read-only without preventing safe formatting', async () => {
    const w = world();
    w.paragraphs[1].getOoxml.mockImplementation(() => { throw new Error('GeneralException'); });
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove blank lines and bold headings' });
    expect(proposal.ops).toEqual([{ font: { bold: false } }]);
    expect(w.deps.log).toHaveBeenCalledWith(expect.stringContaining('will be preserved: GeneralException'), 'warning');
    await applyFormatProposal(w.deps, proposal);
    expect(w.events).toEqual(['bold:false']);
});

test('an uncertain deletion is terminal and restores the original tracking mode', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove blank lines', cleanupOnly: true });
    w.context.sync.mockImplementation(async () => { if (w.events.some((e) => e.startsWith('delete:')) && w.document.changeTrackingMode === 'TrackAll') throw new Error('GeneralException'); });
    const result = await applyFormatProposal(w.deps, proposal);
    expect(result).toMatchObject({ partial: true, deletedParagraphs: 0 });
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
    expect(w.paragraphs[1].delete).toHaveBeenCalledTimes(1);
});

test('cancellation between scan and delete prevents a native write', async () => {
    const w = world(); const controller = new AbortController();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'remove blank lines', cleanupOnly: true });
    w.context.sync.mockImplementation(async () => { controller.abort(); });
    await expect(applyFormatProposal(w.deps, proposal, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(w.events).toEqual([]);
});

test('a forged cleanup op without host evidence cannot delete paragraphs', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'bold headings' });
    proposal.ops.push({ cleanup: { emptyParagraphs: true } });
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/not verified/);
    expect(w.events).toEqual([]);
});
