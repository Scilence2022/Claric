/** @jest-environment jsdom */
jest.mock('../src/lib/llm-client.js', () => ({ sendPrompt: jest.fn(), sendPromptStream: jest.fn() }));
jest.mock('../src/taskpane/word-render-tools.js', () => ({ canReadWordVisuals: jest.fn(() => false), createWordVisualTools: jest.fn() }));
const { sendPrompt } = require('../src/lib/llm-client.js');
const { canReadWordVisuals, createWordVisualTools } = require('../src/taskpane/word-render-tools.js');
const { prepareFormatProposal, applyFormatProposal } = require('../src/taskpane/word-actions.js');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const prose = 'This paragraph explains the observed system behavior and the practical evidence supporting a safe native formatting operation.';
const escapeXml = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const defaults = { alignment: 'Left', lineSpacing: 12, spaceBefore: 0, spaceAfter: 6, leftIndent: 0, rightIndent: 0, firstLineIndent: 0 };

function world(specs = [{}, { text: '6 Discussion' }, { picture: true, text: '' }, { inTable: true }], { ids = false } = {}) {
    const writes = [];
    const control = { idFailure: false, idPending: false };
    const states = specs.map((spec, index) => ({ text: prose, style: '正文', styleBuiltIn: 'Normal',
        id: `session-${index + 1}`, pictureId: '11111111', inTable: false, relation: 'Inside', ...defaults, ...spec }));
    const searches = [];
    const paragraphs = states.map((state, index) => {
        const font = { bold: false, italic: false, underline: 'None', strikeThrough: false, doubleStrikeThrough: false,
            superscript: false, subscript: false, allCaps: false, smallCaps: false, color: '#000000',
            highlightColor: 'None', name: 'Times New Roman', size: 12, load: jest.fn() };
        const xml = () => {
            let textRuns = `<w:r><w:t>${escapeXml(state.text)}</w:t></w:r>`;
            if (state.characterFormat?.bold) {
                const offset = state.text.indexOf(state.characterFormat.match);
                textRuns = `<w:r><w:t>${escapeXml(state.text.slice(0, offset))}</w:t></w:r>`
                    + `<w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(state.characterFormat.match)}</w:t></w:r>`
                    + `<w:r><w:t>${escapeXml(state.text.slice(offset + state.characterFormat.match.length))}</w:t></w:r>`;
            }
            return `<w:p xmlns:w="${W}" xmlns:wp="${WP}" xmlns:wp14="${WP14}">`
                + `<w:pPr><w:pStyle w:val="${escapeXml(state.styleBuiltIn)}"/>${state.directAlignment ? `<w:jc w:val="${state.alignment}"/>` : ''}</w:pPr>`
                + textRuns + (state.picture ? `<w:r><w:drawing><wp:inline wp14:anchorId="${state.pictureId}"/></w:drawing></w:r>` : '') + '</w:p>';
        };
        const search = jest.fn((match) => {
            const matchedFont = {};
            Object.defineProperty(matchedFont, 'bold', {
                get: () => state.characterFormat?.bold || false,
                set: (value) => { writes.push({ index, property: 'bold', value, match }); state.characterFormat = { match, bold: value }; },
            });
            return { load: jest.fn(), items: state.text.includes(match) ? [{ text: match, font: matchedFont }] : [] };
        });
        searches.push(search);
        const paragraph = {
            load: jest.fn((properties) => { if (properties === 'uniqueLocalId') control.idPending = true; }),
            getOoxml: jest.fn(() => ({ value: xml() })),
            getRange: jest.fn(() => ({ font, compareLocationWith: jest.fn(() => ({ value: state.relation })),
                search, paragraphs: { items: [paragraph], load: jest.fn() } })),
            parentTableOrNullObject: { load: jest.fn(), get isNullObject() { return !state.inTable; } },
        };
        ['text', 'style', 'uniqueLocalId'].forEach((property) => Object.defineProperty(paragraph, property,
            { get: () => property === 'uniqueLocalId' ? state.id : state[property] }));
        Object.defineProperty(paragraph, 'styleBuiltIn', { get: () => state.styleBuiltIn,
            set: (value) => { writes.push({ index, property: 'styleBuiltIn', value }); state.styleBuiltIn = value; } });
        Object.keys(defaults).forEach((property) => Object.defineProperty(paragraph, property, {
            get: () => state[property],
            set: (value) => { writes.push({ index, property, value }); if (!state.ignoreWrites) state[property] = value; },
        }));
        return paragraph;
    });
    const range = { isNullObject: false, load: jest.fn(), font: {}, insertBookmark: jest.fn(),
        paragraphs: { items: paragraphs, load: jest.fn() },
        search: jest.fn((match) => ({ load: jest.fn(), items: searches.flatMap((search) => search(match).items) })),
        get text() { return states.map((state) => state.text).join('\r'); },
        getOoxml: jest.fn(() => ({ value: `<w:document xmlns:w="${W}"><w:body>`
            + paragraphs.map((paragraph) => paragraph.getOoxml().value).join('') + '</w:body></w:document>' })),
    };
    const document = { changeTrackingMode: 'TrackMineOnly', load: jest.fn(),
        body: { getRange: jest.fn(() => range) }, getSelection: jest.fn(() => range),
        getBookmarkRangeOrNullObject: jest.fn(() => range), deleteBookmark: jest.fn() };
    if (ids) document.getParagraphByUniqueLocalId = jest.fn();
    const context = { document, sync: jest.fn(async () => {
        if (control.idPending) { control.idPending = false; if (control.idFailure) throw new Error('License does not support local IDs'); }
    }) };
    global.Word = { run: jest.fn(async (fn) => fn(context)),
        Alignment: { left: 'Left', centered: 'Centered', right: 'Right', justified: 'Justified' },
        Style: { normal: 'Normal', heading1: 'Heading1' }, BuiltInStyleName: { normal: 'Normal', heading1: 'Heading1' },
        UnderlineType: { none: 'None', single: 'Single' }, HighlightColor: { none: 'None' },
        ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' } };
    const deps = { appState: { config: { backend: 'mock', providers: { mock: { model: 'm' } }, trackChangesEnabled: true } }, log: jest.fn() };
    return { states, paragraphs, range, document, context, writes, searches, control, deps };
}

beforeEach(() => {
    jest.clearAllMocks(); canReadWordVisuals.mockReturnValue(false);
    sendPrompt.mockResolvedValue('[{"paragraphStyle":"normal","paragraph":{"alignment":"justified"}}]');
});
afterEach(() => { delete global.Word; });

test('explicit body alignment uses native evidence without a model or PDF and excludes non-body Normal paragraphs', async () => {
    const w = world();
    canReadWordVisuals.mockReturnValue(true);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(createWordVisualTools).not.toHaveBeenCalled();
    expect(proposal.ops).toEqual([{ paragraphRole: 'body', paragraph: { alignment: 'justified' } }]);
    expect(proposal.anchor.targetPlan).toMatchObject({ mode: 'targets', bodyOnly: true });
    expect(proposal.targetSummary).toMatchObject({ verifiedParagraphs: 1, excludedParagraphs: 3 });
    expect(await applyFormatProposal(w.deps, proposal))
        .toMatchObject({ applied: true, appliedRanges: 1, verifiedParagraphs: 1, partial: false });
    expect(w.writes).toEqual([{ index: 0, property: 'alignment', value: 'Justified' }]);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    expect(w.document.getSelection).not.toHaveBeenCalled();
    expect(w.range.insertBookmark).not.toHaveBeenCalled();
});

test('document drawing identity drift is preserved and cannot reject an unrelated body-alignment Apply', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    w.states[2].pictureId = '22222222';
    const result = await applyFormatProposal(w.deps, proposal);
    expect(result).toMatchObject({ applied: true, partial: false, verifiedParagraphs: 1 });
    expect(w.states[2].pictureId).toBe('22222222');
    expect(w.writes).toHaveLength(1);
    expect(w.deps.log.mock.calls.some(([message]) => /baseline v3 mismatch/.test(message))).toBe(false);
});

test.each(['text', 'inherited alignment', 'native role'])('target %s drift rejects Apply before native formatting', async (kind) => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    if (kind === 'text') w.states[0].text += ' A new concurrent statement.';
    if (kind === 'inherited alignment') w.states[0].alignment = 'Centered';
    if (kind === 'native role') w.states[0].styleBuiltIn = 'Heading1';
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow();
    expect(w.writes).toEqual([]);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('no verified prose causes the deterministic body request to fail before any model or renderer call', async () => {
    const w = world([{ text: '6 Discussion' }, { text: 'Figure 1. System model.' }, { text: '', picture: true }]);
    await expect(prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' }))
        .rejects.toThrow(/No verified formatting paragraphs matched/);
    expect(sendPrompt).not.toHaveBeenCalled();
    expect(createWordVisualTools).not.toHaveBeenCalled();
    expect(w.writes).toEqual([]);
});

test('generic style selector uses portable built-in Normal rather than its localized display name', async () => {
    const w = world([{}, { style: 'Heading 1', styleBuiltIn: 'Heading1', text: '1 Introduction' }]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Set Normal paragraphs to justified', scope: 'document' });
    expect(sendPrompt).toHaveBeenCalledTimes(1);
    expect(proposal.anchor.targetPlan.entries[0].ids).toEqual(['p1']);
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, verifiedParagraphs: 1 });
    expect(w.states[1].alignment).toBe('Left');
});

test('unknown or removed selector cannot become whole-document formatting at Apply', async () => {
    const w = world();
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    proposal.ops[0] = { paragraphStyle: 'MissingStyle', paragraph: { alignment: 'justified' } };
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/No verified formatting paragraphs matched/);
    expect(w.writes).toEqual([]);
});

test('body-only substring font operation searches only eligible paragraph Content ranges', async () => {
    const w = world([{}, { text: '6 evidence supporting the model' }, { inTable: true }]);
    sendPrompt.mockResolvedValue('[{"match":"evidence","font":{"bold":true}}]');
    const proposal = await prepareFormatProposal(w.deps, { instruction: '将正文中的 evidence 加粗，保留标题和表格。', scope: 'document' });
    expect(proposal.anchor.targetPlan).toMatchObject({ mode: 'strict', bodyOnly: true });
    expect(proposal.anchor.targetPlan.entries[0].ids).toEqual(['p1']);
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, appliedRanges: 1, partial: false });
    expect(w.range.search).not.toHaveBeenCalled();
    expect(w.searches[0]).toHaveBeenCalledWith('evidence', { matchCase: true, matchWholeWord: false });
    expect(w.searches[1]).not.toHaveBeenCalled();
    expect(w.searches[2]).not.toHaveBeenCalled();
    expect(w.writes).toEqual([{ index: 0, property: 'bold', value: true, match: 'evidence' }]);
    expect(w.paragraphs[0].getOoxml().value).toContain('<w:rPr><w:b/></w:rPr>');
    expect(w.states[1].characterFormat).toBeUndefined();
    expect(w.states[2].characterFormat).toBeUndefined();
});

test('a modified valid operation signature cannot bypass the preflight proposal contract', async () => {
    const w = world([{}]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    proposal.ops[0] = { paragraphRole: 'body', paragraph: { alignment: 'right' } };
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/target set changed/);
    expect(w.writes).toEqual([]);
    expect(w.states[0].alignment).toBe('Left');
});

test.each(['heading style', 'list'])('strict %s operations retain the original whole-scope OOXML gate', async (kind) => {
    const w = world();
    const payload = kind === 'heading style' ? { styleBuiltIn: 'heading1' } : { listType: 'bullet' };
    sendPrompt.mockResolvedValue(JSON.stringify([{ paragraphRole: 'body', paragraph: payload }]));
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Adjust the body paragraphs using the requested native format', scope: 'document' });
    expect(proposal.anchor.targetPlan.mode).toBe('strict');
    w.states[2].pictureId = '22222222';
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/anchored formatting baseline changed/);
    expect(w.writes).toEqual([]);
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('strict style assignment uses BuiltInStyleName when Word.Style is a client-object class', async () => {
    const w = world([{}]);
    global.Word.Style = class Style {};
    sendPrompt.mockResolvedValue('[{"paragraphRole":"body","paragraph":{"styleBuiltIn":"heading1"}}]');
    const proposal = await prepareFormatProposal(w.deps, { instruction: 'Format the body paragraphs as Heading 1', scope: 'document' });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, partial: false });
    expect(w.writes).toEqual([{ index: 0, property: 'styleBuiltIn', value: 'Heading1' }]);
    expect(w.states[0].styleBuiltIn).toBe('Heading1');
});

test.each([
    { styleBuiltIn: 'banana' },
    { styleBuiltIn: 'heading1', alignment: 'distribute' },
])('strict native formatting enums are validated before a proposal can claim applicable changes: %j', async (paragraph) => {
    const w = world([{}]);
    sendPrompt.mockResolvedValue(JSON.stringify([{ paragraphRole: 'body', paragraph }]));
    await expect(prepareFormatProposal(w.deps, { instruction: 'Adjust native formatting of the body paragraphs', scope: 'document' }))
        .rejects.toThrow(/Unknown|cannot.*verify|invalid|unsupported/i);
    expect(w.writes).toEqual([]);
});

test('selection uses its original bookmark and excludes partial paragraph boundaries', async () => {
    const w = world([{ relation: 'Contains' }, { relation: 'Inside' }, { relation: 'OverlapsAfter' }]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'selection' });
    const other = { text: w.range.text, load: jest.fn() };
    w.document.getSelection.mockReturnValue(other);
    expect(proposal.targetSummary).toMatchObject({ verifiedParagraphs: 1, excludedParagraphs: 2 });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: true, appliedRanges: 1 });
    expect(w.writes).toEqual([{ index: 1, property: 'alignment', value: 'Justified' }]);
    expect(w.document.getBookmarkRangeOrNullObject).toHaveBeenCalledWith(proposal.anchor.bookmark);
    expect(w.document.body.getRange).not.toHaveBeenCalled();
    expect(other).not.toHaveProperty('alignment');
});

test('missing selection bookmark never falls back to document or a new equal-text selection', async () => {
    const w = world([{}]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'selection' });
    w.range.isNullObject = true;
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/changed or disappeared/);
    expect(w.document.body.getRange).not.toHaveBeenCalled();
    expect(w.writes).toEqual([]);
});

test.each(['expanded', 'narrowed'])('a %s selection target set refuses Apply even when all paragraph IDs and text survive', async (kind) => {
    const w = world([{ relation: 'Contains' }, { relation: 'Inside' }], { ids: true });
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'selection' });
    expect(proposal.anchor.targetPlan.entries[0].ids).toEqual(['session-2']);
    if (kind === 'expanded') w.states[0].relation = 'Inside';
    else w.states[1].relation = 'Contains';
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/target set changed|No verified formatting paragraphs matched/);
    expect(w.writes).toEqual([]);
    expect(w.document.body.getRange).not.toHaveBeenCalled();
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
});

test('read-back rejection reports partial and a single-use proposal cannot replay failed writes', async () => {
    const w = world([{ ignoreWrites: true }]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: false, partial: true });
    expect(w.deps.log).toHaveBeenCalledWith(expect.stringContaining('read-back did not confirm alignment'), 'warning');
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
    expect(w.writes).toHaveLength(1);
});

test('lost native sync after a queued write consumes the proposal and cannot replay the uncertain command', async () => {
    const w = world([{}]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    let failed = false;
    w.context.sync.mockImplementation(async () => {
        if (w.writes.length && !failed) { failed = true; throw new Error('Lost native sync response'); }
    });
    expect(await applyFormatProposal(w.deps, proposal)).toMatchObject({ applied: false, partial: true });
    expect(w.states[0].alignment).toBe('Justified'); // a failed response does not prove the write was rolled back
    expect(w.document.changeTrackingMode).toBe('TrackMineOnly');
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
    expect(w.writes).toHaveLength(1);
});

test('already-justified body paragraphs are verified as no-ops and cannot replay', async () => {
    const w = world([{ alignment: 'Justified' }]);
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    expect(await applyFormatProposal(w.deps, proposal))
        .toMatchObject({ applied: false, partial: false, alreadySatisfied: true, verifiedParagraphs: 1, noopParagraphs: 1 });
    expect(w.writes).toEqual([]);
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/already been attempted/);
});

test('local paragraph identity evidence is not downgraded if Microsoft 365 ID reads stop working', async () => {
    const w = world([{}], { ids: true });
    const proposal = await prepareFormatProposal(w.deps, { instruction: '正文修改为两端对齐', scope: 'document' });
    expect(proposal.anchor.targetPlan.identityMode).toBe('local-id');
    w.control.idFailure = true;
    await expect(applyFormatProposal(w.deps, proposal)).rejects.toThrow(/identity could not be verified/);
    expect(w.writes).toEqual([]);
});
