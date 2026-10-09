/** @jest-environment jsdom */
const { readFormatInventory, compileFormatTargetPlan, validateFormatTargetPlan,
    applyVerifiedFormatTargets } = require('../src/taskpane/format-target-runtime.js');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const prose = 'This paragraph explains how the native document model preserves user content while applying a specific formatting change safely.';
const escapeXml = (text) => String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const alignmentOp = { paragraphRole: 'body', paragraph: { alignment: 'justified' } };
const scalarDefaults = { alignment: 'Left', lineSpacing: 12, spaceBefore: 0, spaceAfter: 6,
    leftIndent: 0, rightIndent: 0, firstLineIndent: 0 };
const fontDefaults = { bold: false, italic: false, underline: 'None', strikeThrough: false,
    doubleStrikeThrough: false, superscript: false, subscript: false, allCaps: false, smallCaps: false,
    color: '#000000', highlightColor: 'None', name: 'Times New Roman', size: 12 };

function world(specs = [{}], { ids = false } = {}) {
    const writes = [];
    const control = { idsFail: false, idsPending: false, capsUnsupported: false, failXml: new Set(), failText: new Set() };
    const states = specs.map((spec, index) => ({ text: prose, style: 'Normal', styleBuiltIn: 'Normal',
        relation: 'Inside', inTable: false, id: `local-${index + 1}`, pictureId: '11111111',
        ...scalarDefaults, ...spec, font: { ...fontDefaults, ...(spec.font || {}) } }));
    const paragraphs = states.map((state, index) => {
        const font = { load: jest.fn((properties) => {
            if (control.capsUnsupported && /allCaps|smallCaps/.test(properties)) throw new Error('WordApiDesktop 1.3 is unavailable');
        }) };
        Object.keys(fontDefaults).forEach((property) => Object.defineProperty(font, property, {
            get: () => {
                if (control.capsUnsupported && ['allCaps', 'smallCaps'].includes(property)) throw new Error('PropertyNotLoaded');
                return state.font[property];
            },
            set: (value) => { writes.push({ index, group: 'font', property, value }); if (!state.ignoreWrites) state.font[property] = value; },
        }));
        const paragraph = { load: jest.fn((properties) => { if (properties === 'uniqueLocalId') control.idsPending = true; }),
            getRange: jest.fn(() => ({ font, compareLocationWith: jest.fn(() => ({ value: state.relation })) })),
            getOoxml: jest.fn(() => {
                if (control.failXml.has(index)) throw new Error('GeneralException');
                const drawing = state.picture ? `<w:r><w:drawing><wp:inline wp14:anchorId="${state.pictureId}"/></w:drawing></w:r>` : '';
                const props = `${state.outline !== undefined ? `<w:outlineLvl w:val="${state.outline}"/>` : ''}`
                    + `<w:pStyle w:val="${escapeXml(state.styleBuiltIn)}"/>`
                    + (state.directAlignment ? `<w:jc w:val="${escapeXml(state.alignment)}"/>` : '');
                const bold = state.font.bold === true ? '<w:rPr><w:b/></w:rPr>' : '';
                return { value: `<w:p xmlns:w="${W}" xmlns:wp="${WP}" xmlns:wp14="${WP14}"><w:pPr>${props}</w:pPr>`
                    + `<w:r>${bold}<w:t>${escapeXml(state.text)}</w:t></w:r>${drawing}</w:p>` };
            }),
            parentTableOrNullObject: { load: jest.fn(), get isNullObject() { return !state.inTable; } },
        };
        ['text', 'style', 'styleBuiltIn', 'uniqueLocalId'].forEach((property) => Object.defineProperty(paragraph, property,
            { get: () => {
                if (property === 'text' && control.failText.has(index)) throw new Error('ItemNotLoaded: text');
                return property === 'uniqueLocalId' ? state.id : state[property];
            } }));
        Object.keys(scalarDefaults).forEach((property) => Object.defineProperty(paragraph, property, {
            get: () => state[property],
            set: (value) => { writes.push({ index, group: 'paragraph', property, value }); if (!state.ignoreWrites) state[property] = value; },
        }));
        return paragraph;
    });
    const scopeRange = { paragraphs: { items: paragraphs, load: jest.fn() } };
    const context = { document: ids ? { getParagraphByUniqueLocalId: jest.fn() } : {}, sync: jest.fn(async () => {
        if (control.idsPending) {
            control.idsPending = false;
            if (control.idsFail) throw new Error('Not supported by this license');
        }
    }) };
    return { context, scopeRange, states, paragraphs, writes, control };
}

beforeEach(() => {
    global.Word = { Alignment: { left: 'Left', centered: 'Centered', right: 'Right', justified: 'Justified' },
        BuiltInStyleName: { normal: 'Normal', heading1: 'Heading1' },
        UnderlineType: { none: 'None', single: 'Single' }, HighlightColor: { none: 'None', yellow: 'Yellow' } };
});
afterEach(() => { delete global.Word; delete global.Office; });

test('captures localized Normal and custom body prose, excluding manual headings, tables and picture paragraphs', async () => {
    const w = world([{ style: '正文' }, { style: 'Academic prose', styleBuiltIn: 'Other' },
        { text: '6 Discussion' }, { inTable: true }, { picture: true }, { text: 'Figure 1. Native document model.' }]);
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(inventory.descriptors.map((d) => d.role)).toEqual(['body', 'body', 'heading', 'table', 'object', 'caption']);
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    expect(plan.entries[0].ids).toEqual(['p1', 'p2']);
    expect(plan.summary).toMatchObject({ verifiedParagraphs: 2, excludedParagraphs: 4 });
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {}))
        .toMatchObject({ applied: true, appliedRanges: 2, verifiedParagraphs: 2 });
    expect(w.writes.map((write) => write.index)).toEqual([0, 1]);
});

test('complete short prose with incorrect bold is not confidently excluded as a manual heading', async () => {
    const w = world([{ font: { bold: true } }, {}, { text: '6 Discussion', font: { bold: true } }]);
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(['body', 'unknown']).toContain(inventory.descriptors[0].role);
    expect(inventory.descriptors[2].role).toBe('heading');
    const plan = compileFormatTargetPlan(inventory, [{ paragraphRole: 'body', font: { bold: false } }], { bodyOnly: true });
    if (inventory.descriptors[0].role === 'body') expect(plan.entries[0].ids).toContain('p1');
    else expect(plan.summary.exclusions).toContainEqual(expect.objectContaining({ id: 'p1', role: 'unknown' }));
});

test('non-target inline drawing attribute churn does not invalidate body alignment', async () => {
    const w = world([{}, { picture: true }]);
    const before = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(before, [alignmentOp], { bodyOnly: true });
    w.states[1].pictureId = '22222222';
    const after = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const current = validateFormatTargetPlan(plan, after, [alignmentOp]);
    await applyVerifiedFormatTargets(w.context, w.scopeRange, current, after, {});
    expect(w.writes).toEqual([{ index: 0, group: 'paragraph', property: 'alignment', value: 'Justified' }]);
    expect(w.states[1].pictureId).toBe('22222222');
});

test.each(['text', 'alignment', 'direct formatting', 'role', 'native style'])('rejects target %s changes before a write', async (kind) => {
    const w = world();
    const before = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(before, [alignmentOp], { bodyOnly: true });
    if (kind === 'text') w.states[0].text += ' A concurrent edit.';
    if (kind === 'alignment') w.states[0].alignment = 'Centered'; // inherited effective value, unchanged XML
    if (kind === 'direct formatting') w.states[0].directAlignment = true;
    if (kind === 'role') w.states[0].inTable = true;
    if (kind === 'native style') w.states[0].styleBuiltIn = 'Heading1';
    const after = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(() => validateFormatTargetPlan(plan, after, [alignmentOp])).toThrow();
    expect(w.writes).toEqual([]);
});

test('overlapping operations compile one final expectation with later values winning', async () => {
    const w = world([{}, { text: `${prose} Additional context.` }]);
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const ops = [alignmentOp, { paragraphIds: ['p2'], paragraph: { alignment: 'right', spaceAfter: 12 } }];
    const plan = compileFormatTargetPlan(inventory, ops);
    expect(plan.final.p2.paragraph).toEqual({ alignment: 'Right', spaceAfter: 12 });
    await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {});
    expect(w.states.map((state) => state.alignment)).toEqual(['Justified', 'Right']);
    expect(w.writes.filter((write) => write.index === 1 && write.property === 'alignment')).toHaveLength(1);
});

test('paragraph font operations use native read-back for each requested property', async () => {
    const w = world();
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const ops = [{ paragraphRole: 'body', font: { italic: true, underline: 'single', size: 13 } }];
    const plan = compileFormatTargetPlan(inventory, ops, { bodyOnly: true });
    expect(plan.mode).toBe('targets');
    const result = await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {});
    expect(result).toMatchObject({ applied: true, verifiedParagraphs: 1 });
    expect(w.states[0].font).toMatchObject({ italic: true, underline: 'Single', size: 13 });
    expect(w.writes.every((write) => write.group === 'font')).toBe(true);
});

test('font names remain distinct when spacing differs instead of being treated as enum aliases', async () => {
    const w = world([{ font: { name: 'Noto Sans' } }]);
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const ops = [{ paragraphRole: 'body', font: { name: 'NotoSans' } }];
    const plan = compileFormatTargetPlan(inventory, ops, { bodyOnly: true });
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {}))
        .toMatchObject({ applied: true, noopParagraphs: 0 });
    expect(w.writes).toEqual([{ index: 0, group: 'font', property: 'name', value: 'NotoSans' }]);
});

test('unavailable desktop-only caps properties cannot block ordinary body alignment', async () => {
    const w = world();
    w.control.capsUnsupported = true;
    global.Office = { context: { requirements: { isSetSupported: jest.fn(() => false) } } };
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    expect(plan.entries[0].ids).toEqual(['p1']);
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {}))
        .toMatchObject({ applied: true, verifiedParagraphs: 1 });
});

test('unavailable desktop-only caps operation fails preflight before any native write', async () => {
    const w = world();
    w.control.capsUnsupported = true;
    global.Office = { context: { requirements: { isSetSupported: jest.fn(() => false) } } };
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const ops = [{ paragraphRole: 'body', font: { allCaps: true } }];
    expect(() => compileFormatTargetPlan(inventory, ops, { bodyOnly: true })).toThrow(/allCaps|unsupported|unavailable|cannot.*verify/i);
    expect(w.writes).toEqual([]);
});

test('highlight color is a native string and does not require a nonexistent Word.HighlightColor enum', async () => {
    const w = world([{ font: { highlightColor: null } }]);
    delete global.Word.HighlightColor;
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const ops = [{ paragraphRole: 'body', font: { highlightColor: '#FFFF00' } }];
    const plan = compileFormatTargetPlan(inventory, ops, { bodyOnly: true });
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {}))
        .toMatchObject({ applied: true, verifiedParagraphs: 1 });
    expect(w.states[0].font.highlightColor).toBe('#FFFF00');
});

test.each([
    ['Yellow', '#ffff00', '#FFFF00'],
    ['#FFFF00', 'Yellow', 'yellow'],
])('equivalent native highlight name and hex values do not create a stale baseline or redundant write', async (initial, current, requested) => {
    const w = world([{ font: { highlightColor: initial } }]);
    delete global.Word.HighlightColor;
    const before = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const ops = [{ paragraphRole: 'body', font: { highlightColor: requested } }];
    const plan = compileFormatTargetPlan(before, ops, { bodyOnly: true });
    w.states[0].font.highlightColor = current;
    const after = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const currentPlan = validateFormatTargetPlan(plan, after, ops);
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, currentPlan, after, {}))
        .toMatchObject({ applied: false, alreadySatisfied: true, verifiedParagraphs: 1 });
    expect(w.writes).toEqual([]);
});

test.each([null, '#FFFF00', ''])('remove-highlight uses native null, preserving the distinction between no highlight and mixed color: %j', async (initial) => {
    const w = world([{ font: { highlightColor: initial } }]);
    delete global.Word.HighlightColor;
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const ops = [{ paragraphRole: 'body', font: { highlightColor: 'none' } }];
    const plan = compileFormatTargetPlan(inventory, ops, { bodyOnly: true });
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {}))
        .toMatchObject({ applied: initial !== null, alreadySatisfied: initial === null, verifiedParagraphs: 1 });
    expect(w.states[0].font.highlightColor).toBeNull();
    expect(w.writes).toHaveLength(initial === null ? 0 : 1);
    if (initial !== null) expect(w.writes[0].value).toBeNull();
});

test('selection Content containment preserves partial boundaries and never expands them', async () => {
    const w = world([{ relation: 'Contains' }, { relation: 'InsideStart' }, { relation: 'OverlapsAfter' }]);
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'selection' });
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    expect(plan.entries[0].ids).toEqual(['p2']);
    await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {});
    expect(w.writes.map((write) => write.index)).toEqual([1]);
    for (const paragraph of w.paragraphs) expect(paragraph.getRange).toHaveBeenCalledWith('Content');
});

test('an unreadable paragraph is preserved while independent body targets remain editable', async () => {
    const w = world([{}, {}]);
    w.control.failXml.add(1);
    const log = jest.fn();
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document', log });
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    expect(plan.entries[0].ids).toEqual(['p1']);
    expect(plan.summary.exclusions[0]).toMatchObject({ id: 'p2' });
    expect(log).toHaveBeenCalledWith(expect.stringContaining('GeneralException'), 'warning');
});

test('unloaded text in failure recovery cannot crash inventory or block other verified body targets', async () => {
    const w = world([{}, {}]);
    w.control.failXml.add(1);
    w.control.failText.add(1);
    const log = jest.fn();
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document', log });
    expect(inventory.descriptors[0]).toMatchObject({ role: 'body', verified: true });
    expect(inventory.descriptors[1].verified).toBe(false);
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    expect(plan.entries[0].ids).toEqual(['p1']);
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, {}))
        .toMatchObject({ applied: true, verifiedParagraphs: 1 });
    expect(w.writes).toEqual([{ index: 0, group: 'paragraph', property: 'alignment', value: 'Justified' }]);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('GeneralException'), 'warning');
});

test.each([
    { paragraphRole: 'body', paragraph: { alignment: 'distribute' } },
    { paragraphStyle: 'Heading9', paragraph: { alignment: 'justified' } },
    { paragraphIds: ['p999'], paragraph: { alignment: 'justified' } },
])('invalid or unmatched selector fails without native writes: %j', async (op) => {
    const w = world();
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(() => compileFormatTargetPlan(inventory, [op])).toThrow();
    expect(w.writes).toEqual([]);
});

test('already satisfied paragraphs are read back and consume a verified no-op', async () => {
    const w = world([{ alignment: 'Justified' }]);
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    const anchor = {};
    expect(await applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, anchor))
        .toEqual({ applied: false, appliedRanges: 0, verifiedParagraphs: 1, noopParagraphs: 1, alreadySatisfied: true });
    expect(anchor.attempted).toBe(true);
    expect(w.writes).toEqual([]);
    expect(w.paragraphs[0].load).toHaveBeenCalledWith('text,alignment');
});

test('native setter acceptance cannot substitute for read-back confirmation', async () => {
    const w = world([{ ignoreWrites: true }]);
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    const anchor = {};
    await expect(applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, anchor))
        .rejects.toThrow(/read-back did not confirm alignment/);
    expect(anchor.attempted).toBe(true);
    expect(w.states[0].alignment).toBe('Left');
});

test('read-back detects a concurrent content edit after native formatting sync', async () => {
    const w = world();
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    const anchor = {};
    w.context.sync.mockImplementationOnce(async () => { w.states[0].text += ' New content during Apply.'; });
    await expect(applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, anchor))
        .rejects.toThrow(/read-back detected changed paragraph text/);
    expect(anchor.attempted).toBe(true);
});

test('pre-aborted target Apply never queues a native write or consumes the proposal', async () => {
    const w = world();
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(inventory, [alignmentOp], { bodyOnly: true });
    const anchor = {};
    const controller = new AbortController(); controller.abort();
    await expect(applyVerifiedFormatTargets(w.context, w.scopeRange, plan, inventory, anchor, { signal: controller.signal }))
        .rejects.toMatchObject({ name: 'AbortError' });
    expect(w.writes).toEqual([]);
    expect(anchor.attempted).toBeUndefined();
});

test('local ID read failure cannot silently downgrade a captured local-ID plan', async () => {
    const w = world([{}], { ids: true });
    const before = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(before.identityMode).toBe('local-id');
    const plan = compileFormatTargetPlan(before, [alignmentOp], { bodyOnly: true });
    w.control.idsFail = true;
    const after = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(after.identityMode).toBe('ordinal');
    expect(() => validateFormatTargetPlan(plan, after, [alignmentOp])).toThrow(/identity could not be verified/);
    expect(w.writes).toEqual([]);
});

test('a new Word session cannot replace native target identities with equal text', async () => {
    const w = world([{}], { ids: true });
    const before = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(before, [alignmentOp], { bodyOnly: true });
    w.states[0].id = 'another-session';
    const after = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(() => validateFormatTargetPlan(plan, after, [alignmentOp])).toThrow(/target set changed/);
});

test('ordinal fallback rejects changed paragraph order instead of searching equal text', async () => {
    const w = world([{}, { text: `${prose} Second paragraph.` }]);
    const before = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    const plan = compileFormatTargetPlan(before, [alignmentOp], { bodyOnly: true });
    [w.states[0].text, w.states[1].text] = [w.states[1].text, w.states[0].text];
    const after = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(() => validateFormatTargetPlan(plan, after, [alignmentOp])).toThrow(/order or scope changed/);
});

test.each([
    { insert: { text: 'Heading', position: 'start' } },
    { cleanup: { emptyParagraphs: true } },
    { paragraphRole: 'body', paragraph: { listType: 'bullet' } },
    { paragraphRole: 'body', paragraph: { styleBuiltIn: 'heading1' } },
    { match: 'paragraph', font: { bold: true } },
])('structural and character-range operations retain strict mode: %j', async (op) => {
    const w = world();
    const inventory = await readFormatInventory(w.context, w.scopeRange, { scope: 'document' });
    expect(compileFormatTargetPlan(inventory, [op]).mode).toBe('strict');
});
