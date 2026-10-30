/** @jest-environment jsdom */
jest.mock('../src/lib/llm-client.js', () => ({ sendPrompt: jest.fn(), sendPromptStream: jest.fn() }));
jest.mock('../src/taskpane/word-render-tools.js', () => ({ canReadWordVisuals: () => false }));
const { sendPrompt } = require('../src/lib/llm-client.js');
const { prepareFormatProposal, applyFormatProposal } = require('../src/taskpane/word-actions.js');
const { parseFormatOps, buildFormatPrompt, describeFormatOp } = require('../src/lib/format-ops.js');
const { capabilityPrompt } = require('../src/lib/capability-catalog.js');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const prose = 'The current paragraph contains reviewed scientific prose and explains the observed results in detail.';
const xml = (text) => `<w:p xmlns:w="${W}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const numbered = (id, from, to) => ({ paragraphIds: [id], headingNumber: { from, to } });

function world(specs = [{}]) {
    const writes = [];
    const states = specs.map((spec) => ({ text: prose, currentText: prose, style: 'Body Text', styleBuiltIn: 'BodyText',
        isListItem: false, alignment: 'Left', inTable: false, relation: 'Inside', prefixLocation: 'Equal', ...spec }));
    const paragraphs = states.map((state, index) => {
        const font = { load: jest.fn(), bold: false };
        const range = (kind) => ({ font, getReviewedText: () => ({ get value() { return state.currentText; } }),
            compareLocationWith: () => ({ value: state.relation }),
            search: (text) => ({ load: jest.fn(), items: [{ text,
                getRange: () => ({ compareLocationWith: () => ({ value: state.prefixLocation }) }),
                insertText: (replacement, position) => {
                    writes.push({ index, replacement, position });
                    if (!state.ignoreTextWrite) state.currentText = replacement + state.currentText.slice(text.length);
                    // Native text can include both old and inserted text under tracking.
                    state.text = text + state.currentText;
                } }] }), kind });
        const p = { load: jest.fn(), getRange: range,
            parentTableOrNullObject: { load: jest.fn(), get isNullObject() { return !state.inTable; } },
            getOoxml: () => ({ value: state.ooxml || xml(state.currentText) }) };
        for (const key of ['text', 'style', 'styleBuiltIn', 'isListItem']) Object.defineProperty(p, key, { get: () => state[key] });
        Object.defineProperty(p, 'alignment', { get: () => state.alignment,
            set: (value) => { writes.push({ index, alignment: value }); state.alignment = value; } });
        return p;
    });
    const scope = { load: jest.fn(), isNullObject: false, insertBookmark: jest.fn(),
        get text() { return states.map((s) => s.text).join('\r'); },
        paragraphs: { load: jest.fn(), items: paragraphs },
        getOoxml: () => ({ value: `<w:body xmlns:w="${W}">${states.map((s) => xml(s.currentText)).join('')}</w:body>` }) };
    const document = { changeTrackingMode: 'TrackMineOnly', load: jest.fn(), body: { getRange: () => scope },
        getSelection: () => scope, getBookmarkRangeOrNullObject: () => scope, deleteBookmark: jest.fn() };
    const context = { document, sync: jest.fn(async () => {}) };
    global.Word = { run: async (callback) => callback(context), Alignment: { justified: 'Justified', left: 'Left' },
        ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' } };
    const deps = { appState: { config: { backend: 'mock', providers: { mock: { model: 'test' } }, trackChangesEnabled: true } }, log: jest.fn() };
    return { states, writes, deps, document };
}
afterEach(() => { delete global.Word; jest.clearAllMocks(); });

test('native paragraph alignment includes revised prose and table cells exported with neighboring paragraphs', async () => {
    const revised = `<w:p xmlns:w="${W}"><w:del><w:r><w:delText>Old text</w:delText></w:r></w:del><w:ins><w:r><w:t>${prose}</w:t></w:r></w:ins></w:p>`;
    const expanded = `<w:body xmlns:w="${W}"><w:tbl><w:tr><w:tc>${xml(prose)}</w:tc><w:tc>${xml('Neighbor')}</w:tc></w:tr></w:tbl></w:body>`;
    const w = world([{ text: `Old text${prose}`, ooxml: revised }, { ooxml: expanded, inTable: true }]);
    sendPrompt.mockResolvedValue('[{"paragraph":{"alignment":"justified"}}]');
    const proposal = await prepareFormatProposal(w.deps, { instruction: '修改所有段落为两端对齐', scope: 'document' });
    expect(proposal.targetSummary).toMatchObject({ verifiedParagraphs: 2, excludedParagraphs: 0 });
    // Render/export churn must not invalidate native property setters.
    w.states[0].ooxml = xml(prose);
    w.states[1].ooxml = xml(prose);
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, verifiedParagraphs: 2, partial: false });
    expect(w.writes).toEqual([{ index: 0, alignment: 'Justified' }, { index: 1, alignment: 'Justified' }]);
    expect(w.states[0].text).toBe(`Old text${prose}`);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('body-only native alignment includes revised prose but preserves headings and table cells', async () => {
    const w = world([{ ooxml: `<w:p xmlns:w="${W}"><w:del><w:r><w:delText>Old</w:delText></w:r></w:del><w:r><w:t>${prose}</w:t></w:r></w:p>` },
        { text: '2 Methods', currentText: '2 Methods', styleBuiltIn: 'Heading1' }, { inTable: true }]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文段落两端对齐', scope: 'document' });
    expect(proposal.targetSummary.verifiedParagraphs).toBe(1);
    await applyFormatProposal(w.deps, proposal);
    expect(w.writes).toEqual([{ index: 0, alignment: 'Justified' }]);
});

test.each(['text', 'currentText', 'styleBuiltIn', 'alignment', 'inTable', 'scope'])('native alignment still rejects actual %s drift', async (kind) => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文段落两端对齐', scope: 'selection' });
    if (kind === 'scope') w.states[0].relation = 'Contains';
    else w.states[0][kind] = kind === 'inTable' ? true : 'Changed';
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/changed|matched/);
    expect(w.writes).toEqual([]);
});

const heading = { text: '4.1 Model definition', currentText: '4.1 Model definition', style: 'Heading 2', styleBuiltIn: 'Heading2' };
test('literal heading numbers are corrected as prefix-only tracked edits with reviewed read-back', async () => {
    const w = world([heading, { ...heading, text: '4.4 Results', currentText: '4.4 Results' }]);
    sendPrompt.mockResolvedValue(JSON.stringify([numbered('p1', '4.1', '2.1'), numbered('p2', '4.4', '3.1')]));
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Correct the literal heading numbering', scope: 'document' });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, appliedRanges: 2, partial: false });
    expect(w.writes).toEqual([{ index: 0, replacement: '2.1', position: 'Replace' }, { index: 1, replacement: '3.1', position: 'Replace' }]);
    expect(w.states.map((s) => s.currentText)).toEqual(['2.1 Model definition', '3.1 Results']);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
});

test.each([{ isListItem: true }, { isListItem: undefined }, { styleBuiltIn: 'BodyText', currentText: prose },
    { text: '4.10 Model definition', currentText: '4.10 Model definition' }])('heading renumbering rejects unproven literal headings: %j', async (change) => {
    const w = world([{ ...heading, ...change }]);
    sendPrompt.mockResolvedValue(JSON.stringify([numbered('p1', '4.1', '2.1')]));
    await expect(prepareFormatProposal(w.deps, { instruction: 'Correct heading numbers', scope: 'document' })).rejects.toThrow(/literal leading number/);
    expect(w.writes).toEqual([]);
});

test('all heading prefixes are located before any heading is edited', async () => {
    const w = world([heading, { ...heading, prefixLocation: 'After' }]);
    sendPrompt.mockResolvedValue(JSON.stringify([numbered('p1', '4.1', '2.1'), numbered('p2', '4.1', '2.2')]));
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Correct heading numbers', scope: 'document' });
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/paragraph start/);
    expect(w.writes).toEqual([]);
});

test('generated field numbers cannot masquerade as literal heading prefixes', async () => {
    const ooxml = `<w:p xmlns:w="${W}"><w:fldSimple w:instr="SEQ section"><w:r><w:t>4.1</w:t></w:r></w:fldSimple><w:r><w:t> Model definition</w:t></w:r></w:p>`;
    const w = world([{ ...heading, ooxml }]);
    sendPrompt.mockResolvedValue(JSON.stringify([numbered('p1', '4.1', '2.1')]));
    await expect(prepareFormatProposal(w.deps, { instruction: 'Correct heading numbers', scope: 'document' })).rejects.toThrow(/literal leading number/);
    expect(w.writes).toEqual([]);
});

test('a heading switched to automatic numbering after review is rejected before a write', async () => {
    const w = world([heading]);
    sendPrompt.mockResolvedValue(JSON.stringify([numbered('p1', '4.1', '2.1')]));
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Correct heading numbers', scope: 'document' });
    w.states[0].isListItem = true;
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/literal leading number/);
    expect(w.writes).toEqual([]);
});

test('failed heading read-back reports partial and prevents replay', async () => {
    const w = world([{ ...heading, ignoreTextWrite: true }]);
    sendPrompt.mockResolvedValue(JSON.stringify([numbered('p1', '4.1', '2.1')]));
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Correct heading numbers', scope: 'document' });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: false, partial: true });
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
});

test('the contract routes literal numbering to native inspection without trusting prior list speculation', () => {
    const op = numbered('p1', '4.1', '2.1');
    expect(parseFormatOps(JSON.stringify([op]))).toEqual([op]);
    expect(describeFormatOp(op)).toContain('4.1 → 2.1');
    for (const invalid of [{ ...op, paragraphIds: ['p1', 'p2'] }, { headingNumber: op.headingNumber },
        { ...op, headingNumber: { from: '4.1', to: 'New title' } }, { ...op, font: { bold: true } }]) {
        expect(parseFormatOps(JSON.stringify([invalid]))).toEqual([]);
    }
    expect(capabilityPrompt()).toContain('Never assume dotted heading numbers');
    expect(buildFormatPrompt('Fix headings', '', 'document', [{ id: 'p1', ...heading, isListItem: false }])).toContain('"isListItem":false');
});
