/**
 * word-actions.js keeps same-named wrappers for the table write paths that
 * live in the lazily loaded word-table-actions.js chunk. These tests pin the
 * delegation contract: table code downloads only for confirmed table work,
 * a failed download never touches Word, and loaded actions receive the
 * original request, the private LLM helper, and the caller's cancellation.
 */

jest.mock('../src/taskpane/task-module-loader.js', () => ({ loadTableActions: jest.fn() }));
jest.mock('../src/lib/llm-client.js', () => ({
    sendPrompt: jest.fn(async () => 'Revised clause'),
    sendPromptStream: jest.fn(),
    stripMarkdown: jest.fn((t) => t),
    stripChunkDelimiters: jest.fn((t) => t),
}));

const { loadTableActions } = require('../src/taskpane/task-module-loader.js');
const {
    prepareTableProposal, applyTableProposal, prepareSelectionAmendment, applySelectionAmendment,
} = require('../src/taskpane/word-actions.js');

function makeDeps() {
    return {
        appState: {
            config: {
                backend: 'mock',
                providers: { mock: { url: '', apiKey: '', model: 'mock-model', apiPath: '' } },
                trackChangesEnabled: true,
            },
            promptManager: {
                getActivePrompt: () => null,
                composeMessages: jest.fn((text) => [{ role: 'user', content: text }]),
            },
            platform: 'PC',
        },
        log: jest.fn(),
    };
}

/** Word.run over a plain text selection (no table anywhere). */
function mockPlainSelection(text = 'Selected clause') {
    const absent = { isNullObject: true, load: jest.fn() };
    const selection = {
        text, load: jest.fn(),
        parentTableOrNullObject: absent, parentTableCellOrNullObject: absent,
        paragraphs: { items: [], load: jest.fn() },
    };
    const context = { document: { getSelection: () => selection }, sync: jest.fn(async () => {}) };
    global.Word = { run: jest.fn(async (fn) => fn(context)) };
}

/** Word.run over a multi-cell selection R1C1 → R2C2 of a uniform table. */
function mockTableSelection() {
    const table = {
        isNullObject: false, rowCount: 2, isUniform: true, load: jest.fn(),
        values: [['Header A', 'Header B'], ['a', 'b']],
    };
    const cellAt = (rowIndex, cellIndex) => ({ isNullObject: false, rowIndex, cellIndex, load: jest.fn() });
    const selection = {
        parentTableOrNullObject: table,
        parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
        getRange: (loc) => ({ parentTableCellOrNullObject: loc === 'Start' ? cellAt(0, 0) : cellAt(1, 1) }),
    };
    const context = { document: { getSelection: () => selection }, sync: jest.fn(async () => {}) };
    global.Word = { run: jest.fn(async (fn) => fn(context)), RangeLocation: { start: 'Start', end: 'End' } };
}

function nativeTableActions() {
    return {
        prepareTableAmendment: jest.fn(async () => ({ tablePatch: { cells: [] }, tableItems: [] })),
        applyTablePatch: jest.fn(async () => ({ cellsApplied: 1, cellsSkipped: 0, rowOpsApplied: 0, warnings: [] })),
        applyMixedTableAmendment: jest.fn(async () => true),
        prepareTableProposal: jest.fn(async () => ({ instruction: '插入一个三行三列的表格', spec: {}, model: null, warnings: [] })),
        applyTableProposal: jest.fn(async () => ({ inserted: true, rowCount: 3, columnCount: 3, tracked: true, warnings: [] })),
    };
}

beforeEach(() => { jest.clearAllMocks(); global.Word = { run: jest.fn() }; });
afterEach(() => { delete global.Word; });

test('failed table download cannot read or write Word', async () => {
    const failure = new Error('Table tools unavailable');
    loadTableActions.mockRejectedValue(failure);
    const deps = makeDeps();
    await expect(prepareTableProposal(deps, { instruction: '插入一个三行三列的表格' })).rejects.toBe(failure);
    await expect(applyTableProposal(deps, { spec: { rows: [['a']] } })).rejects.toBe(failure);
    await expect(applySelectionAmendment(deps, { tablePatch: { cells: [], rowOps: [] } })).rejects.toBe(failure);
    await expect(applySelectionAmendment(deps, { mixedTable: true, selectionText: 'a', amendedText: 'b' })).rejects.toBe(failure);
    expect(Word.run).not.toHaveBeenCalled();
    expect(deps.log).not.toHaveBeenCalledWith(expect.stringContaining('Applying changes'), expect.anything());
});

test('table creation wrappers invoke loaded actions once, with the request, LLM helper and cancellation', async () => {
    const native = nativeTableActions();
    loadTableActions.mockResolvedValue(native);
    const deps = makeDeps();
    const signal = new AbortController().signal;
    const args = { instruction: '插入一个三行三列的表格', onToken: jest.fn(), signal };
    const proposal = await prepareTableProposal(deps, args);
    expect(native.prepareTableProposal).toHaveBeenCalledTimes(1);
    expect(native.prepareTableProposal).toHaveBeenCalledWith(
        expect.objectContaining({ appState: deps.appState, log: deps.log, sendActionRequest: expect.any(Function) }), args);
    expect(await applyTableProposal(deps, proposal)).toEqual(expect.objectContaining({ inserted: true }));
    expect(native.applyTableProposal).toHaveBeenCalledWith(deps, proposal);
    expect(loadTableActions.mock.calls).toEqual([[{ signal }], []]);
});

test('table and mixed selection proposals apply through the loaded module', async () => {
    const native = nativeTableActions();
    loadTableActions.mockResolvedValue(native);
    const deps = makeDeps();
    const patchProposal = { tablePatch: { cells: [{ row: 2, col: 1, text: 'A' }], rowOps: [] } };
    expect(await applySelectionAmendment(deps, patchProposal)).toEqual(expect.objectContaining({ cellsApplied: 1 }));
    expect(native.applyTablePatch).toHaveBeenCalledWith(deps, patchProposal);

    const mixedProposal = { mixedTable: true, selectionText: 'Caption\na', amendedText: 'Caption\nA' };
    expect(await applySelectionAmendment(deps, mixedProposal)).toBeUndefined();
    expect(native.applyMixedTableAmendment).toHaveBeenCalledWith(deps, mixedProposal);
    expect(deps.log).toHaveBeenCalledWith('Changes applied successfully', 'success');
    expect(Word.run).not.toHaveBeenCalled();
});

test('a confirmed multi-cell region downloads the patch protocol with the caller signal', async () => {
    mockTableSelection();
    const native = nativeTableActions();
    loadTableActions.mockResolvedValue(native);
    const deps = makeDeps();
    const signal = new AbortController().signal;
    const result = await prepareSelectionAmendment(deps, { promptTemplate: 'Capitalize {selection}', signal });
    expect(result).toEqual(expect.objectContaining({ tablePatch: { cells: [] } }));
    expect(loadTableActions).toHaveBeenCalledWith({ signal });
    const [tableDeps, tableArgs] = native.prepareTableAmendment.mock.calls[0];
    expect(tableDeps).toEqual(expect.objectContaining({ log: deps.log, sendActionRequest: expect.any(Function) }));
    expect(tableArgs).toEqual(expect.objectContaining({
        promptTemplate: 'Capitalize {selection}', signal,
        backendConfig: expect.objectContaining({ model: 'mock-model' }),
        tableRegion: expect.objectContaining({ rowCount: 2, colCount: 2, bounds: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 } }),
    }));
});

test('plain-text selection edits never download table code', async () => {
    mockPlainSelection();
    const deps = makeDeps();
    const proposal = await prepareSelectionAmendment(deps, { promptTemplate: 'Polish {selection}' });
    expect(proposal).toEqual(expect.objectContaining({ selectionText: 'Selected clause', amendedText: 'Revised clause' }));
    expect(proposal.tablePatch).toBeUndefined();
    expect(proposal.mixedTable).toBeUndefined();

    // A stale flat apply stops before any write — still without table code.
    mockPlainSelection('Different selection');
    expect(await applySelectionAmendment(deps, proposal)).toEqual(expect.objectContaining({ skipped: true }));
    expect(loadTableActions).not.toHaveBeenCalled();
});
