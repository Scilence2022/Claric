/** @jest-environment jsdom */
const { preparePageLayoutProposal: prepare, applyPageLayoutProposal: apply, discardPageLayoutProposal: discard,
    readPageLayout: read } = require('../src/taskpane/page-layout-actions.js');
const { PAGE_NUMBERS, PAGE_BOOLEANS, PAGE_ENUMS } = require('../src/lib/page-layout.js');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function world(count = 3) {
    const writes = [], control = { selected: [1], ignoreWrites: false, onSync: null };
    const state = Array.from({ length: count }, (_, i) => ({
        text: `Section ${i + 1} body`, protectedForForms: false,
        pageSetup: { ...Object.fromEntries(PAGE_NUMBERS.map((key) => [key, 0])),
            ...Object.fromEntries(PAGE_BOOLEANS.map((key) => [key, false])),
            ...Object.fromEntries(Object.entries(PAGE_ENUMS).map(([key, values]) => [key, values[0]])),
            pageWidth: 612, pageHeight: 792, topMargin: 72, bottomMargin: 72, leftMargin: 72, rightMargin: 72,
            headerDistance: 36, footerDistance: 36, paperSize: 'Letter', orientation: 'Portrait' },
        lineNumbering: { isActive: true, countBy: 1, startingNumber: 1, distanceFromText: 12, restartMode: 'RestartContinuous' },
        columns: { evenlySpaced: true, lineBetween: false, flowDirection: 'LeftToRight', items: [{ width: 468, spaceAfter: 0 }] },
    }));
    function mutate(i, group, patch) {
        writes.push({ i, group, patch });
        if (!control.ignoreWrites) Object.assign(state[i][group], patch);
    }
    const sections = state.map((s, i) => {
        const columns = {
            load: jest.fn(),
            getIsEvenlySpaced: () => ({ value: s.columns.evenlySpaced }),
            getHasLineBetween: () => ({ value: s.columns.lineBetween }),
            getFlowDirection: () => ({ value: s.columns.flowDirection }),
            setCount: (n) => mutate(i, 'columns', { items: Array.from({ length: n }, () => ({ width: 180, spaceAfter: 18 })) }),
            setIsEvenlySpaced: (evenlySpaced) => mutate(i, 'columns', { evenlySpaced }),
            setHasLineBetween: (lineBetween) => mutate(i, 'columns', { lineBetween }),
            setFlowDirection: (flowDirection) => mutate(i, 'columns', { flowDirection }),
            get items() { return s.columns.items; },
        };
        return { load: jest.fn(), get protectedForForms() { return s.protectedForForms; },
            pageSetup: { ...s.pageSetup, load: jest.fn(),
                set: (patch) => { mutate(i, 'pageSetup', patch); Object.assign(sections[i].pageSetup, s.pageSetup); },
                lineNumbering: { ...s.lineNumbering, load: jest.fn(),
                    set: (patch) => { mutate(i, 'lineNumbering', patch); Object.assign(sections[i].pageSetup.lineNumbering, s.lineNumbering); } },
                textColumns: columns },
            body: { getRange: () => ({ index: i, compareLocationWith: (target) => ({ value: target.index === i ? 'Equal' : 'Before' }) }),
                getOoxml: () => ({ value: `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>${s.text}</w:t></w:r></w:p>`
                    + `<w:sectPr><w:lnNumType w:countBy="${s.lineNumbering.countBy}"/></w:sectPr></w:body></w:document>` }) },
        };
    });
    const collection = { load: jest.fn(), items: sections };
    const context = { document: { sections: collection, getSelection: () => ({ sections: { load: jest.fn(), items: control.selected.map((i) => sections[i]) } }) },
        sync: jest.fn(async () => { control.onSync?.(); }) };
    global.Word = { run: jest.fn(async (fn) => fn(context)) };
    global.Office = { context: { requirements: { isSetSupported: jest.fn(() => true) } } };
    const deps = { appState: { config: { backend: 'mock', providers: { mock: { model: 'test-model' } } } },
        log: jest.fn(), sendActionRequest: jest.fn(async () => '[]') };
    const plan = (patch, ids = ['s1']) => deps.sendActionRequest.mockResolvedValue(JSON.stringify([{ sectionIds: ids, patch }]));
    return { state, sections, writes, control, context, deps, plan };
}
afterEach(() => { delete global.Word; delete global.Office; });

test('removes line numbers from every native section without a model or body write', async () => {
    const f = world();
    const proposal = await prepare(f.deps, { instruction: '去掉文档中的行号' });
    expect(proposal.ops.map((o) => o.sectionId)).toEqual(['s1', 's2', 's3']);
    expect(f.writes).toEqual([]);
    expect(f.deps.sendActionRequest).not.toHaveBeenCalled();
    expect(Office.context.requirements.isSetSupported).toHaveBeenCalledWith('WordApiDesktop', '1.3');
    expect(await apply(f.deps, proposal)).toMatchObject({ applied: true, verified: true, verifiedSections: 3 });
    expect(f.writes).toEqual([0, 1, 2].map((i) => ({ i, group: 'lineNumbering', patch: { isActive: false } })));
    expect(f.state.map((s) => s.text)).toEqual(['Section 1 body', 'Section 2 body', 'Section 3 body']);
    await expect(apply(f.deps, proposal)).rejects.toThrow('already attempted');
});

test('selection pins containing sections and does not follow a later caret move', async () => {
    const f = world();
    f.control.selected = [1, 2];
    const proposal = await prepare(f.deps, { instruction: 'Remove line numbers', scope: 'selection' });
    expect(proposal.ops.map((o) => o.sectionId)).toEqual(['s2', 's3']);
    f.control.selected = [0];
    await apply(f.deps, { ...proposal, ops: proposal.ops.slice(0, 1) });
    expect(f.writes).toEqual([{ i: 1, group: 'lineNumbering', patch: { isActive: false } }]);
});

test('native read exposes section settings, never content fingerprints or writes', async () => {
    const f = world();
    const records = await read(f.deps, { scope: 'selection' });
    expect(records).toHaveLength(1);
    expect(records[0]).toMatchObject({ id: 's2', lineNumbering: { isActive: true } });
    expect(records[0].bodyFingerprint).toBeUndefined();
    expect(records[0].selected).toBeUndefined();
    expect(f.writes).toEqual([]);
    expect(f.deps.sendActionRequest).not.toHaveBeenCalled();
});

test('uses the strict model contract for compound geometry and numbering settings', async () => {
    const f = world();
    const patch = { pageSetup: { leftMargin: 56.69, rightMargin: 56.69, orientation: 'Landscape',
        gutter: 10, mirrorMargins: true, headerDistance: 30, footerDistance: 30,
        sectionStart: 'NewPage', differentFirstPageHeaderFooter: true },
    lineNumbering: { isActive: true, countBy: 5, startingNumber: 10, restartMode: 'RestartPage', distanceFromText: 15 },
    columns: { count: 2, evenlySpaced: true, lineBetween: true, flowDirection: 'LeftToRight' } };
    f.plan(patch);
    const proposal = await prepare(f.deps, { instruction: 'Use landscape, 2 cm margins and two columns, restart line numbers per page' });
    expect(f.deps.sendActionRequest).toHaveBeenCalledWith(f.deps, expect.objectContaining({ model: 'test-model' }),
        expect.stringContaining('OBSERVED SECTION SETTINGS'), expect.any(Object));
    expect(f.deps.sendActionRequest.mock.calls[0][2]).not.toContain('Section 1 body');
    expect(f.writes).toEqual([]);
    await apply(f.deps, proposal);
    expect(f.state[0]).toMatchObject({ pageSetup: patch.pageSetup, lineNumbering: patch.lineNumbering });
    expect(f.state[0].columns).toMatchObject({ evenlySpaced: true, lineBetween: true });
    expect(f.state[0].columns.items).toHaveLength(2);
});

test('custom column dimensions apply after the native column count is refreshed', async () => {
    const f = world();
    const columns = { count: 2, evenlySpaced: false, lineBetween: true,
        items: [{ width: 200, spaceAfter: 18 }, { width: 230, spaceAfter: 0 }] };
    f.plan({ columns });
    const proposal = await prepare(f.deps, { instruction: 'Set custom column widths' });
    await apply(f.deps, proposal);
    expect(f.state[0].columns.items).toEqual(columns.items);
});

test('already satisfied deterministic requests avoid empty proposals and writes', async () => {
    const f = world(1);
    f.state[0].lineNumbering.isActive = false;
    f.sections[0].pageSetup.lineNumbering.isActive = false;
    const proposal = await prepare(f.deps, { instruction: 'Remove line numbers' });
    expect(proposal).toMatchObject({ ops: [], alreadySatisfied: true });
    expect(f.deps.sendActionRequest).not.toHaveBeenCalled();
    await expect(apply(f.deps, proposal)).rejects.toThrow('unavailable');
    expect(f.writes).toEqual([]);
});

test.each(['body', 'layout', 'section-count', 'protection'])('rejects a stale %s before any setter', async (change) => {
    const f = world();
    const proposal = await prepare(f.deps, { instruction: 'Remove line numbers' });
    if (change === 'body') f.state[0].text += ' edited';
    if (change === 'layout') f.sections[0].pageSetup.leftMargin = 50;
    if (change === 'section-count') f.sections.pop();
    if (change === 'protection') f.state[0].protectedForForms = true;
    await expect(apply(f.deps, proposal)).rejects.toThrow('changed after preparation');
    expect(f.writes).toEqual([]);
});

test('forged anchors, altered patches, duplicate ops and discarded proposals cannot write', async () => {
    const f = world();
    const proposal = await prepare(f.deps, { instruction: 'Remove line numbers' });
    await expect(apply(f.deps, { ...proposal, anchor: {} })).rejects.toThrow('unavailable');
    await expect(apply(f.deps, { ...proposal, ops: [{ sectionId: 's1', patch: { pageSetup: { leftMargin: 20 } } }] })).rejects.toThrow('changed after review');
    await expect(apply(f.deps, { ...proposal, ops: [proposal.ops[0], proposal.ops[0]] })).rejects.toThrow('changed after review');
    await discard(f.deps, proposal);
    await expect(apply(f.deps, proposal)).rejects.toThrow('unavailable');
    await discard(f.deps, { anchor: {} });
    expect(f.writes).toEqual([]);
});

test.each([
    { pageSetup: { leftMargin: 400, rightMargin: 400 } },
    { pageSetup: { topMargin: 400, bottomMargin: 400 } },
    { pageSetup: { gutter: 700, gutterPosition: 'Top' } },
    { columns: { count: 2, evenlySpaced: false, items: [{ width: 400 }, { width: 400 }] } },
    { columns: { evenlySpaced: false, items: [{ width: 100 }, { width: 100 }] } },
])('rejects impossible geometry %j before staging', async (patch) => {
    const f = world(); f.plan(patch);
    await expect(prepare(f.deps, { instruction: 'Set layout' })).rejects.toThrow();
    expect(f.writes).toEqual([]);
});

test('unknown/out-of-scope native targets and protected sections fail before writes', async () => {
    const f = world();
    f.plan({ pageSetup: { leftMargin: 36 } }, ['s1']);
    await expect(prepare(f.deps, { instruction: 'Set left margin', scope: 'selection' })).rejects.toThrow('out-of-scope');
    f.state[1].protectedForForms = true;
    await expect(prepare(f.deps, { instruction: 'Remove line numbers' })).rejects.toThrow('protected');
    expect(f.writes).toEqual([]);
});

test('unsupported host exits before Word reads or model calls', async () => {
    const f = world(); Office.context.requirements.isSetSupported.mockReturnValue(false);
    await expect(prepare(f.deps, { instruction: 'Remove line numbers' })).rejects.toThrow('WordApiDesktop 1.3');
    await expect(read(f.deps)).rejects.toThrow('WordApiDesktop 1.3');
    await expect(apply(f.deps, {})).rejects.toThrow('WordApiDesktop 1.3');
    expect(Word.run).not.toHaveBeenCalled();
    expect(f.deps.sendActionRequest).not.toHaveBeenCalled();
});

test('pre-apply cancellation has no writes; cancellation after a setter reports partial application', async () => {
    const f = world(); const controller = new AbortController(); controller.abort();
    await expect(prepare(f.deps, { instruction: 'Remove line numbers', signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    const proposal = await prepare(f.deps, { instruction: 'Remove line numbers' });
    await expect(apply(f.deps, proposal, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    expect(f.writes).toEqual([]);
    const second = new AbortController();
    f.control.onSync = () => { if (f.writes.length) second.abort(); };
    await expect(apply(f.deps, proposal, { signal: second.signal })).rejects.toMatchObject({ mutationAttempted: true });
    const writes = f.writes.length;
    await expect(apply(f.deps, proposal)).rejects.toThrow('already attempted');
    expect(f.writes).toHaveLength(writes);
});

test.each(['sync', 'readback', 'body', 'other-section'])('reports uncertain mutation on %s failure, never retries the native action', async (failure) => {
    const f = world();
    const proposal = await prepare(f.deps, { instruction: 'Remove line numbers' });
    const selected = { ...proposal, ops: [proposal.ops[0]] };
    if (failure === 'readback') f.control.ignoreWrites = true;
    f.control.onSync = () => {
        if (!f.writes.length) return;
        if (failure === 'sync') throw new Error('Host sync failed');
        if (failure === 'body') f.state[0].text = 'Unexpected modification';
        if (failure === 'other-section') f.sections[1].pageSetup.topMargin = 10;
    };
    await expect(apply(f.deps, selected)).rejects.toMatchObject({ mutationAttempted: true, message: expect.stringContaining('may be partially applied') });
    expect(f.writes).toHaveLength(1);
    await expect(apply(f.deps, selected)).rejects.toThrow('already attempted');
    expect(f.writes).toHaveLength(1);
});

test('invalid read scope, empty sections, missing API data and ambiguous selections fail closed', async () => {
    const f = world();
    await expect(read(f.deps, { scope: 'paragraph' })).rejects.toThrow('scope');
    await expect(prepare(f.deps, { scope: 'paragraph' })).rejects.toThrow('scope');
    f.control.selected = [1, 1];
    await expect(read(f.deps, { scope: 'selection' })).rejects.toThrow('identify');
    f.sections[0].pageSetup.leftMargin = NaN;
    await expect(read(f.deps)).rejects.toThrow('could not read');
    f.sections[0].pageSetup.leftMargin = 72;
    f.state[0].columns.evenlySpaced = undefined;
    await expect(read(f.deps)).rejects.toThrow('column settings');
    f.sections.length = 0;
    await expect(read(f.deps)).rejects.toThrow('1–100');
    expect(f.writes).toEqual([]);
});
