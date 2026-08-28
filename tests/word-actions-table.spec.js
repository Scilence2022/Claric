/**
 * Unit tests for the table-selection amendment route in
 * src/taskpane/word-actions.js (multi-cell patch protocol, plan A + B).
 *
 * Covers:
 * - prepareSelectionAmendment detects a multi-cell selection and returns a
 *   parsed tablePatch + review items instead of flat amended text
 * - applySelectionAmendment applies cell patches before row ops, in
 *   descending row order
 * - Platform split for row ops: desktop runs them tracked, Word for the web
 *   runs them untracked with a warning
 * - Stale-table guards (selection left the table, row count drifted)
 */

// --- Mock the strategy/LLM layer (module-level side-effect free imports) ---

jest.mock('../src/lib/llm-client.js', () => ({
  sendPrompt: jest.fn(),
  sendPromptStream: jest.fn(),
  stripMarkdown: jest.fn((t) => t),
}));

jest.mock('../src/lib/word-diff/index.js', () => ({
  applyTokenMapStrategy: jest.fn().mockResolvedValue(undefined),
  applySentenceDiffStrategy: jest.fn().mockResolvedValue(undefined),
}));

jest.mock('../src/lib/word-diff/char-diff.js', () => ({
  hasCjk: jest.fn(() => false),
  applyCharDiffStrategy: jest.fn().mockResolvedValue(undefined),
  computeCharEdits: jest.fn(() => []),
  sliceSearchPieces: jest.fn((t) => [t]),
}));

const { sendPrompt, sendPromptStream } = require('../src/lib/llm-client.js');
const { applyTokenMapStrategy } = require('../src/lib/word-diff/index.js');
const {
  prepareSelectionAmendment,
  applySelectionAmendment,
  readSelectionTableRegion,
} = require('../src/taskpane/word-actions.js');

// --- Mock helpers ---

const TABLE_VALUES = [
  ['Header A', 'Header B'],
  ['old a', 'b'],
  ['c', 'd'],
];

function makeDeps(platform = 'PC') {
  return {
    appState: {
      config: {
        backend: 'mock',
        providers: { mock: { url: '', apiKey: '', model: 'mock-model', apiPath: '' } },
        trackChangesEnabled: true,
      },
      promptManager: {
        getActivePrompt: () => null,
        composeMessages: jest.fn((selectionText) => [{ role: 'user', content: selectionText }]),
        composeMergedMessages: jest.fn((selectionText) => [{ role: 'user', content: selectionText }]),
      },
      platform,
      supportsComments: false,
    },
    log: jest.fn(),
  };
}

/** Word.run mock where the selection covers a multi-cell region. */
function makePrepareContext() {
  const table = {
    isNullObject: false,
    rowCount: 3,
    values: TABLE_VALUES,
    isUniform: true,
    load: jest.fn(),
  };
  const selection = {
    parentTableOrNullObject: table,
    parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
    getRange: jest.fn((loc) => ({
      parentTableCellOrNullObject: loc === 'Start'
        ? { isNullObject: false, rowIndex: 0, cellIndex: 0, load: jest.fn() }
        : { isNullObject: false, rowIndex: 2, cellIndex: 1, load: jest.fn() },
    })),
  };
  return {
    document: { getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
}

/**
 * Word.run mock for a MERGED table: row 1 has a horizontal merge across both
 * columns (values reports the text in the anchor slot, '' in the shadow).
 * getCell(r, c) resolves shadow coordinates to the anchor cell — the exact
 * Word behavior the merge detection relies on.
 */
function makeMergedPrepareContext() {
  const values = [
    ['Wide header', ''],
    ['old a', 'b'],
    ['c', 'd'],
  ];
  // Anchor coordinates per grid slot; (0,1) is covered by the merge at (0,0).
  const anchors = [
    [{ rowIndex: 0, cellIndex: 0 }, { rowIndex: 0, cellIndex: 0 }],
    [{ rowIndex: 1, cellIndex: 0 }, { rowIndex: 1, cellIndex: 1 }],
    [{ rowIndex: 2, cellIndex: 0 }, { rowIndex: 2, cellIndex: 1 }],
  ];
  const cellProxies = {};
  const table = {
    isNullObject: false,
    rowCount: 3,
    values,
    isUniform: false,
    load: jest.fn(),
    getCell: jest.fn((r, c) => {
      const key = `${r},${c}`;
      if (!cellProxies[key]) {
        cellProxies[key] = { ...anchors[r][c], load: jest.fn() };
      }
      return cellProxies[key];
    }),
  };
  const selection = {
    parentTableOrNullObject: table,
    parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
    getRange: jest.fn((loc) => ({
      parentTableCellOrNullObject: loc === 'Start'
        ? { isNullObject: false, rowIndex: 0, cellIndex: 0, load: jest.fn() }
        : { isNullObject: false, rowIndex: 2, cellIndex: 1, load: jest.fn() },
    })),
  };
  return {
    document: { getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
}

/**
 * Word.run mock for the apply phase. Records the order of cell/row writes and
 * every changeTrackingMode assignment.
 */
function makeApplyContext({ tableRowCount = 3 } = {}) {
  const calls = [];
  const trackingModes = [];

  const rows = {};
  const cells = {};
  for (let r = 0; r < 3; r++) {
    rows[r] = {
      delete: jest.fn(() => calls.push(`delete:${r + 1}`)),
      insertRows: jest.fn((loc, count, values) => calls.push(`insert:${r + 1}:${loc}:${JSON.stringify(values)}`)),
      load: jest.fn(),
    };
    for (let c = 0; c < 2; c++) {
      const paraRange = {
        text: TABLE_VALUES[r][c],
        load: jest.fn(),
        insertText: jest.fn(() => calls.push(`cellText:${r + 1},${c + 1}`)),
      };
      cells[`${r},${c}`] = {
        body: {
          paragraphs: {
            items: [{
              getRange: jest.fn(() => paraRange),
              clear: jest.fn(() => calls.push(`clear:${r + 1},${c + 1}`)),
              load: jest.fn(),
            }],
            load: jest.fn(),
          },
        },
        parentRow: rows[r],
        merge: jest.fn((other) => calls.push(`merge:${r + 1},${c + 1}->${other}`)),
      };
    }
  }

  const table = {
    isNullObject: false,
    rowCount: tableRowCount,
    values: TABLE_VALUES,
    isUniform: true,
    load: jest.fn(),
    getCell: jest.fn((r, c) => cells[`${r},${c}`]),
    mergeCells: jest.fn(() => calls.push('merge')),
  };
  const selection = { parentTableOrNullObject: table };

  const context = {
    document: { getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
  let mode = null;
  Object.defineProperty(context.document, 'changeTrackingMode', {
    get: () => mode,
    set: (v) => { mode = v; trackingModes.push(v); },
  });

  return { context, calls, trackingModes, rows, cells, table };
}

function setWordRun(context) {
  global.Word = {
    run: jest.fn((cb) => cb(context)),
    RangeLocation: { start: 'Start', end: 'End', content: 'Content' },
    InsertLocation: { replace: 'Replace', after: 'After', before: 'Before' },
    ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
  };
}

// --- Tests ---

describe('readSelectionTableRegion', () => {
  beforeEach(() => jest.clearAllMocks());

  test('returns null when the selection is not in a table', async () => {
    const context = {
      document: {
        getSelection: () => ({
          parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
          parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
        }),
      },
      sync: jest.fn().mockResolvedValue(undefined),
    };
    setWordRun(context);
    expect(await readSelectionTableRegion(makeDeps())).toBeNull();
  });

  test('returns null when the selection sits inside a single cell', async () => {
    const context = {
      document: {
        getSelection: () => ({
          parentTableOrNullObject: { isNullObject: false, load: jest.fn() },
          parentTableCellOrNullObject: { isNullObject: false, load: jest.fn() },
        }),
      },
      sync: jest.fn().mockResolvedValue(undefined),
    };
    setWordRun(context);
    expect(await readSelectionTableRegion(makeDeps())).toBeNull();
  });

  test('extracts the covered rectangle with 1-based absolute coordinates', async () => {
    setWordRun(makePrepareContext());
    const region = await readSelectionTableRegion(makeDeps());

    expect(region.rowCount).toBe(3);
    expect(region.colCount).toBe(2);
    expect(region.cells).toHaveLength(6);
    expect(region.cells[0]).toEqual({ row: 1, col: 1, text: 'Header A' });
    expect(region.cells[5]).toEqual({ row: 3, col: 2, text: 'd' });
  });

  test('whole-table selection (both endpoints on boundaries) clamps to the whole table', async () => {
    // Whole-table selection via the move handle / Select Table: both
    // endpoint ranges sit exactly on table boundaries and resolve OUTSIDE
    // (null cells) even though containment (parentTable + no anchor cell)
    // proves an in-table multi-cell selection.
    const context = makePrepareContext();
    context.document.getSelection().getRange = jest.fn(() => ({
      parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
    }));
    setWordRun(context);

    const deps = makeDeps();
    const region = await readSelectionTableRegion(deps);

    expect(region).not.toBeNull();
    expect(region.cells).toHaveLength(6);
    expect(region.bounds).toEqual({ startRow: 1, endRow: 3, startCol: 1, endCol: 2 });
    expect(deps.log).toHaveBeenCalledWith(
      expect.stringContaining('table boundary'), 'warning');
  });

  test('selection from the table start through a resolved end keeps the end corner', async () => {
    // Start point on the first-cell boundary (null), end point resolvable:
    // the region runs from the table's first cell to the resolved corner.
    const context = makePrepareContext();
    context.document.getSelection().getRange = jest.fn((loc) => ({
      parentTableCellOrNullObject: loc === 'Start'
        ? { isNullObject: true, load: jest.fn() }
        : { isNullObject: false, rowIndex: 1, cellIndex: 1, load: jest.fn() },
    }));
    setWordRun(context);

    const region = await readSelectionTableRegion(makeDeps());

    expect(region).not.toBeNull();
    expect(region.bounds).toEqual({ startRow: 1, endRow: 2, startCol: 1, endCol: 2 });
    expect(region.cells.map((c) => `${c.row},${c.col}`)).toEqual(['1,1', '1,2', '2,1', '2,2']);
  });

  test('merged table: detects the merge shadow instead of erroring', async () => {
    setWordRun(makeMergedPrepareContext());
    const deps = makeDeps();
    const region = await readSelectionTableRegion(deps);

    expect(region).not.toBeNull();
    expect(region.merged).toBe(true);
    expect(region.shadowKeys).toEqual(new Set(['1,2']));
    // Anchor slot stays editable, shadow slot is flagged.
    expect(region.cells[0]).toEqual({ row: 1, col: 1, text: 'Wide header' });
    expect(region.cells[1]).toEqual({ row: 1, col: 2, text: '', merged: true });
    expect(deps.log).toHaveBeenCalledWith(expect.stringContaining('merged cells'), 'warning');
  });

  test('merged table with failing probes degrades to unknown layout (no crash)', async () => {
    // Hosts that throw ItemNotFound for getCell on merge-covered coordinates:
    // the probe sync is the 3rd sync in the read.
    const context = makeMergedPrepareContext();
    let syncCount = 0;
    context.sync = jest.fn(() => {
        syncCount++;
        return syncCount === 3
            ? Promise.reject(Object.assign(new Error('ItemNotFound'), { name: 'ItemNotFound' }))
            : Promise.resolve(undefined);
    });
    setWordRun(context);

    const deps = makeDeps();
    const region = await readSelectionTableRegion(deps);

    expect(region).not.toBeNull();
    expect(region.merged).toBe(true);
    expect(region.mergedUnknown).toBe(true);
    expect(region.shadowKeys).toEqual(new Set());
    // No cell is flagged read-only — validation moves to apply time.
    expect(region.cells.every((c) => !c.merged)).toBe(true);
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/warning:Merge layout could not be probed/);
  });

  test('unknown merge layout adds the apply-time note to the single-shot prompt', async () => {
    const context = makeMergedPrepareContext();
    let syncCount = 0;
    context.sync = jest.fn(() => {
        syncCount++;
        return syncCount === 3
            ? Promise.reject(Object.assign(new Error('ItemNotFound'), { name: 'ItemNotFound' }))
            : Promise.resolve(undefined);
    });
    setWordRun(context);
    sendPrompt.mockResolvedValue('{"cells":[{"row":2,"col":1,"text":"new a"}]}');

    await prepareSelectionAmendment(makeDeps(), { promptTemplate: 'Fix' });

    const promptText = sendPrompt.mock.calls[0][1];
    expect(promptText).toContain('layout could not be mapped');
    expect(promptText).toContain('skipped at apply time');
  });
});

describe('prepareSelectionAmendment (table route)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('parses the model JSON into a tablePatch + review items', async () => {
    setWordRun(makePrepareContext());
    sendPrompt.mockResolvedValue(
      '{"cells":[{"row":2,"col":1,"text":"new a"},{"row":99,"col":1,"text":"x"}],' +
      '"rowOps":[{"op":"insertAfter","row":1,"values":["n1","n2"]},{"op":"delete","row":3}]}'
    );

    const proposal = await prepareSelectionAmendment(makeDeps(), { promptTemplate: 'Fix the numbers in {selection}' });

    // Prompt carries the grid and the JSON rules, not the flat selection
    const promptText = sendPrompt.mock.calls[0][1];
    expect(promptText).toContain('[R2C1] old a');
    expect(promptText).toContain('"cells"');
    expect(promptText).toContain('the selected table cells');

    expect(proposal.amendedText).toBeNull();
    expect(proposal.tablePatch.cells).toEqual([{ row: 2, col: 1, text: 'new a' }]);
    // Row ops are pre-sorted into descending application order
    expect(proposal.tablePatch.rowOps).toEqual([
      { op: 'delete', row: 3 },
      { op: 'insertAfter', row: 1, values: ['n1', 'n2'] },
    ]);

    // Review items: one per changed cell + one per row op
    expect(proposal.tableItems).toHaveLength(3);
    expect(proposal.tableItems[0]).toMatchObject({ label: 'Cell R2C1', before: 'old a', after: 'new a' });
    expect(proposal.tableItems[1]).toMatchObject({ label: 'Delete row 3', before: 'c | d', after: '' });
    expect(proposal.tableItems[2]).toMatchObject({ label: 'Insert row after row 1', after: 'n1 | n2' });
  });

  test('warns on out-of-bounds coordinates', async () => {
    setWordRun(makePrepareContext());
    sendPrompt.mockResolvedValue('{"cells":[{"row":99,"col":1,"text":"x"}]}');
    const deps = makeDeps();
    await prepareSelectionAmendment(deps, { promptTemplate: 'Fix' });
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/warning:Table patch: Cell coordinate out of bounds/);
  });

  test('throws when the model does not return a patch', async () => {
    setWordRun(makePrepareContext());
    sendPrompt.mockResolvedValue('I rewrote your table as markdown: | a | b |');
    await expect(prepareSelectionAmendment(makeDeps(), { promptTemplate: 'Fix' }))
      .rejects.toThrow(/no JSON object/);
  });

  test('streams via sendPromptStream when token handlers are given', async () => {
    setWordRun(makePrepareContext());
    sendPromptStream.mockResolvedValue('{"cells":[{"row":2,"col":2,"text":"b2"}]}');
    const proposal = await prepareSelectionAmendment(makeDeps(), {
      promptTemplate: 'Fix', onToken: jest.fn(),
    });
    expect(sendPromptStream).toHaveBeenCalled();
    expect(proposal.tablePatch.cells).toEqual([{ row: 2, col: 2, text: 'b2' }]);
  });

  test('merged table: shadow edits dropped, row ops disabled, anchors kept', async () => {
    setWordRun(makeMergedPrepareContext());
    // The model tries to edit the shadow slot AND insert a row — both must
    // be rejected by the protocol; the anchor edit survives.
    sendPrompt.mockResolvedValue(
      '{"cells":[{"row":1,"col":2,"text":"shadow write"},' +
      '{"row":1,"col":1,"text":"Wider header"}],' +
      '"rowOps":[{"op":"insertAfter","row":1,"values":["n1","n2"]}]}'
    );

    const deps = makeDeps();
    const proposal = await prepareSelectionAmendment(deps, { promptTemplate: 'Fix the header' });

    // Prompt carries the merged markers and the merged rules
    const promptText = sendPrompt.mock.calls[0][1];
    expect(promptText).toContain('[R1C2] (merged — read-only)');
    expect(promptText).toContain('MERGED CELLS');

    expect(proposal.tablePatch.cells).toEqual([{ row: 1, col: 1, text: 'Wider header' }]);
    expect(proposal.tablePatch.rowOps).toEqual([]);

    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/warning:Table patch: Cell R1C2 is covered by a merged cell/);
    expect(logged).toMatch(/warning:Table patch: Row operations are not allowed/);
  });
});

describe('applySelectionAmendment (table route)', () => {
  beforeEach(() => jest.clearAllMocks());

  const proposal = {
    selectionText: 'x',
    amendedText: null,
    commentText: null,
    tablePatch: {
      rowCount: 3,
      colCount: 2,
      cells: [{ row: 2, col: 1, text: 'new a' }],
      rowOps: [
        { op: 'delete', row: 3 },
        { op: 'insertAfter', row: 1, values: ['n1', 'n2'] },
      ],
      bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
      originals: TABLE_VALUES,
    },
    tableItems: [],
  };

  test('applies cell patches before row ops, row ops descending', async () => {
    const { context, calls } = makeApplyContext();
    setWordRun(context);
    await applySelectionAmendment(makeDeps('PC'), proposal);

    // Cell write ran through the granular strategy; row ops came after it,
    // delete(row 3) before insertAfter(row 1).
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(applyTokenMapStrategy.mock.calls[0][2]).toBe('old a');
    expect(applyTokenMapStrategy.mock.calls[0][3]).toBe('new a');
    expect(applyTokenMapStrategy.mock.calls[0][5]).toEqual({ trackChanges: false });
    expect(calls).toEqual([
      'delete:3',
      'insert:1:After:[["n1","n2"]]',
    ]);
  });

  test('applies a cell merge: clears away-cells then table.mergeCells', async () => {
    const mergeProposal = {
      selectionText: 'x',
      amendedText: null,
      commentText: null,
      tablePatch: {
        rowCount: 3,
        colCount: 2,
        cells: [],
        rowOps: [],
        merges: [{ op: 'merge', startRow: 2, startCol: 1, endRow: 3, endCol: 2 }],
        bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
        originals: TABLE_VALUES,
      },
      tableItems: [],
    };

    const { context, calls, table } = makeApplyContext();
    setWordRun(context);
    const deps = makeDeps('PC');
    const result = await applySelectionAmendment(deps, mergeProposal);

    // Non-anchor cells (R2C2, R3C1, R3C2) cleared; anchor (R2C1) kept and
    // merged with each other cell in the block via TableCell.merge(mergeTo).
    expect(calls.slice(0, 3)).toEqual(['clear:2,2', 'clear:3,1', 'clear:3,2']);
    const anchor = table.getCell(1, 0);
    expect(anchor.merge).toHaveBeenCalledTimes(3);
    expect(result.mergesApplied).toBe(1);
    // Structural — merge is not tracked as a revision (tracking off).
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toContain('not be tracked');
  });

  test('desktop (PC) keeps tracking on for row ops', async () => {
    const { context, trackingModes } = makeApplyContext();
    setWordRun(context);
    await applySelectionAmendment(makeDeps('PC'), proposal);
    expect(trackingModes).toEqual(['TrackAll', 'TrackAll', 'Off']);
  });

  test('Word for the web applies row ops untracked and warns', async () => {
    const { context, calls, trackingModes } = makeApplyContext();
    setWordRun(context);
    const deps = makeDeps('OfficeOnline');
    await applySelectionAmendment(deps, proposal);

    // Cell phase tracked, row phase forced off, final restore off.
    expect(trackingModes).toEqual(['TrackAll', 'Off', 'Off']);
    expect(calls).toEqual(['delete:3', 'insert:1:After:[["n1","n2"]]']);
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/cannot be tracked as revisions/);
  });

  test('track-changes off applies everything untracked', async () => {
    const { context, trackingModes } = makeApplyContext();
    setWordRun(context);
    const deps = makeDeps('PC');
    deps.appState.config.trackChangesEnabled = false;
    await applySelectionAmendment(deps, proposal);
    expect(trackingModes).toEqual(['Off', 'Off', 'Off']);
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
  });

  test('skips cell writes when the cell text already matches', async () => {
    const { context, calls } = makeApplyContext();
    setWordRun(context);
    const noOpProposal = {
      ...proposal,
      tablePatch: { ...proposal.tablePatch, cells: [{ row: 2, col: 1, text: 'old a' }], rowOps: [] },
    };
    await applySelectionAmendment(makeDeps('PC'), noOpProposal);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    expect(calls).toEqual([]);
  });

  test('throws when the selection left the table', async () => {
    const { context, table } = makeApplyContext();
    table.isNullObject = true;
    setWordRun(context);
    await expect(applySelectionAmendment(makeDeps('PC'), proposal))
      .rejects.toThrow(/no longer inside the table/);
  });

  test('throws when the row count drifted since staging', async () => {
    const { context } = makeApplyContext({ tableRowCount: 4 });
    setWordRun(context);
    await expect(applySelectionAmendment(makeDeps('PC'), proposal))
      .rejects.toThrow(/Table changed/);
  });

  test('merged table applies anchor-cell edits without the old uniform guard', async () => {
    // A merged-table proposal: one anchor cell edit, no row ops. The old
    // code threw on isUniform === false at apply time; now anchor writes go
    // through getCell (which resolves to the anchor) and nothing throws.
    const { context, calls } = makeApplyContext();
    context.document.getSelection().parentTableOrNullObject.isUniform = false;
    const mergedValues = [
      ['Wide header', ''],
      ['old a', 'b'],
      ['c', 'd'],
    ];
    context.document.getSelection().parentTableOrNullObject.values = mergedValues;
    setWordRun(context);

    const mergedProposal = {
      ...proposal,
      tablePatch: {
        ...proposal.tablePatch,
        values: undefined,
        cells: [{ row: 2, col: 1, text: 'new a' }],
        rowOps: [],
        originals: mergedValues,
      },
    };
    await applySelectionAmendment(makeDeps('PC'), mergedProposal);

    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(calls).toEqual([]);
  });

  test('a cell that cannot be addressed (merge-covered) is skipped, not fatal', async () => {
    const { context, table } = makeApplyContext();
    // R1C2 is merge-covered on this host: getCell throws ItemNotFound.
    const realGetCell = table.getCell;
    table.getCell = jest.fn((r, c) => {
        if (r === 0 && c === 1) {
            throw Object.assign(new Error('ItemNotFound'), { name: 'ItemNotFound' });
        }
        return realGetCell(r, c);
    });
    setWordRun(context);

    const deps = makeDeps('PC');
    await applySelectionAmendment(deps, {
      ...proposal,
      tablePatch: {
        ...proposal.tablePatch,
        cells: [
          { row: 1, col: 2, text: 'shadow write' }, // merge-covered → skipped
          { row: 2, col: 1, text: 'new a' },        // good → applied
        ],
        rowOps: [],
      },
    });

    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/warning:Cell R1C2: could not be addressed/);
  });
});

// --- Mixed selection (paragraphs + table) route ---

/**
 * Word.run mock for a selection that overlaps a table without being inside
 * it: caption + one cell + a blank spacer + note. parentTableOrNullObject on
 * the selection itself is null; per-paragraph parents mark cell membership.
 */
function makeMixedPrepareContext() {
  const paras = [
    { text: 'Caption A', inTable: false },
    { text: 'Cell A1', inTable: true },
    { text: '', inTable: false }, // blank spacer — must be filtered out
    { text: 'Note', inTable: false },
  ].map(({ text, inTable }) => ({
    text,
    parentTableOrNullObject: { isNullObject: !inTable, load: jest.fn() },
    load: jest.fn(),
  }));
  const selection = {
    parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
    parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
    paragraphs: { items: paras, load: jest.fn() },
  };
  return {
    document: { getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
}

/**
 * Word.run mock for the mixed apply phase. Each paragraph mock records
 * content-range text loads, deletes and insertParagraph calls; the context
 * records every changeTrackingMode assignment.
 */
function makeMixedApplyContext(spec = [
  { text: 'Caption A', inTable: false },
  { text: 'Cell A1', inTable: true },
  { text: 'Note', inTable: false },
]) {
  const calls = [];
  const trackingModes = [];
  const paraRanges = [];

  const paras = spec.map(({ text, inTable }, i) => {
    const paraRange = {
      text,
      load: jest.fn(),
      insertText: jest.fn((t, loc) => calls.push(`paraText:${i}:${loc}:${t}`)),
    };
    paraRanges.push(paraRange);
    return {
      text,
      parentTableOrNullObject: { isNullObject: !inTable, load: jest.fn() },
      getRange: jest.fn(() => paraRange),
      delete: jest.fn(() => calls.push(`delete:${i}`)),
      insertParagraph: jest.fn((t, loc) => calls.push(`insertPara:${i}:${loc}:${t}`)),
      load: jest.fn(),
    };
  });

  const selection = {
    paragraphs: { items: paras, load: jest.fn() },
  };
  const context = {
    document: { getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
  let mode = null;
  Object.defineProperty(context.document, 'changeTrackingMode', {
    get: () => mode,
    set: (v) => { mode = v; trackingModes.push(v); },
  });

  return { context, calls, trackingModes, paras, paraRanges };
}

const mixedProposal = (amendedText) => ({
  selectionText: 'Caption A\nCell A1\nNote',
  amendedText,
  commentText: null,
  mixedTable: true,
});

describe('prepareSelectionAmendment (mixed table route)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('uses per-paragraph text, flags the proposal, and appends the table note', async () => {
    setWordRun(makeMixedPrepareContext());
    sendPrompt.mockResolvedValue('Caption A\nCell A1 edited\nNote');

    const proposal = await prepareSelectionAmendment(makeDeps(), { promptTemplate: 'Fix {selection}' });

    // Prompt carries non-blank paragraphs one per line, never selection.text
    const promptText = sendPrompt.mock.calls[0][1];
    expect(promptText).toContain('Caption A\nCell A1\nNote');
    expect(promptText).toContain('NOTE: The selection contains a Word table');

    expect(proposal.selectionText).toBe('Caption A\nCell A1\nNote');
    expect(proposal.mixedTable).toBe(true);
    expect(proposal.amendedText).toBe('Caption A\nCell A1 edited\nNote');
  });

  test('pure text selections do not take the mixed route', async () => {
    const context = {
      document: {
        getSelection: () => ({
          text: 'Just text',
          load: jest.fn(),
          parentTableOrNullObject: { isNullObject: true, load: jest.fn() },
          parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
          paragraphs: {
            items: [
              { text: 'Just text', parentTableOrNullObject: { isNullObject: true, load: jest.fn() }, load: jest.fn() },
            ],
            load: jest.fn(),
          },
        }),
      },
      sync: jest.fn().mockResolvedValue(undefined),
    };
    setWordRun(context);
    sendPrompt.mockResolvedValue('Just text edited');

    const proposal = await prepareSelectionAmendment(makeDeps(), { promptTemplate: 'Fix' });
    expect(proposal.mixedTable).toBeUndefined();
  });
});

describe('applySelectionAmendment (mixed table route)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('keep edits diff paragraph content ranges; tracking wraps the writes', async () => {
    const { context, trackingModes } = makeMixedApplyContext();
    setWordRun(context);
    await applySelectionAmendment(makeDeps('PC'), mixedProposal('Caption A\nCell A1 edited\nNote'));

    // Only the cell paragraph changed: one granular diff against its content
    // range, tracked at document level (strategy receives trackChanges:false).
    expect(applyTokenMapStrategy).toHaveBeenCalledTimes(1);
    expect(applyTokenMapStrategy.mock.calls[0][2]).toBe('Cell A1');
    expect(applyTokenMapStrategy.mock.calls[0][3]).toBe('Cell A1 edited');
    expect(applyTokenMapStrategy.mock.calls[0][5]).toEqual({ trackChanges: false });
    expect(trackingModes).toEqual(['TrackAll', 'Off']);
  });

  test('deletes/inserts anchored inside the table are skipped, outside are applied', async () => {
    const { context, calls, paras } = makeMixedApplyContext();
    setWordRun(context);
    const deps = makeDeps('PC');
    await applySelectionAmendment(deps, mixedProposal('New closing line'));

    // Table cell paragraph is never deleted or used as an insert anchor.
    expect(paras[1].delete).not.toHaveBeenCalled();
    expect(paras[1].insertParagraph).not.toHaveBeenCalled();
    expect(paras[0].insertParagraph).toHaveBeenCalledWith('New closing line', 'After');
    expect(calls).toContain('insertPara:0:After:New closing line');

    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/inside a table/);
  });

  test('truncated model output is refused', async () => {
    const { context } = makeMixedApplyContext();
    setWordRun(context);
    await expect(applySelectionAmendment(makeDeps('PC'), mixedProposal('Cap')))
      .rejects.toThrow(/truncated/);
    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
  });

  test('identical output applies nothing', async () => {
    const { context, trackingModes } = makeMixedApplyContext();
    setWordRun(context);
    const deps = makeDeps('PC');
    await applySelectionAmendment(deps, mixedProposal('Caption A\nCell A1\nNote'));

    expect(applyTokenMapStrategy).not.toHaveBeenCalled();
    expect(trackingModes).toEqual([]);
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/No differences found/);
  });
});

// --- Table style ops (read snapshot + apply phases) ---

/**
 * Word.run mock for style-op apply: records property writes via calls, with
 * Table/TableRow border objects, row fonts, and layout method spies.
 */
function makeStyleApplyContext() {
  const calls = [];
  // Border records: plain objects handed out by getBorder(); assertions read
  // their type/color/width after apply.
  const tableBorders = {};
  const rowBorders = {};

  const rows = {};
  const cells = {};
  for (let r = 0; r < 3; r++) {
    rows[r] = {
      font: { bold: null, italic: null, size: null, name: null, color: null },
      shadingColor: null,
      horizontalAlignment: null,
      verticalAlignment: null,
      getBorder: jest.fn((loc) => {
        if (!rowBorders[`${r}:${loc}`]) rowBorders[`${r}:${loc}`] = { type: null, color: null, width: null };
        return rowBorders[`${r}:${loc}`];
      }),
      delete: jest.fn(() => calls.push(`delete:${r + 1}`)),
      load: jest.fn(),
    };
    for (let c = 0; c < 2; c++) {
      cells[`${r},${c}`] = {
        shadingColor: null,
        horizontalAlignment: null,
        verticalAlignment: null,
        columnWidth: null,
        body: { paragraphs: { items: [], load: jest.fn() } },
        parentRow: rows[r],
        load: jest.fn(),
      };
    }
  }

  const table = {
    isNullObject: false,
    rowCount: 3,
    values: TABLE_VALUES,
    isUniform: true,
    load: jest.fn(),
    getCell: jest.fn((r, c) => cells[`${r},${c}`]),
    styleBuiltIn: null,
    style: null,
    styleBandedRows: null,
    styleBandedColumns: null,
    styleFirstColumn: null,
    styleLastColumn: null,
    styleTotalRow: null,
    alignment: null,
    horizontalAlignment: null,
    verticalAlignment: null,
    shadingColor: null,
    headerRowCount: 0,
    width: null,
    font: { bold: null, italic: null, size: null, name: null, color: null },
    getBorder: jest.fn((loc) => {
      if (!tableBorders[loc]) tableBorders[loc] = { type: null, color: null, width: null };
      return tableBorders[loc];
    }),
    autoFitWindow: jest.fn(() => calls.push('autoFitWindow')),
    distributeColumns: jest.fn(() => calls.push('distributeColumns')),
    setCellPadding: jest.fn((side, pt) => calls.push(`padding:${side}:${pt}`)),
  };
  const selection = { parentTableOrNullObject: table };
  const context = {
    document: { getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
  let mode = null;
  const trackingModes = [];
  Object.defineProperty(context.document, 'changeTrackingMode', {
    get: () => mode,
    set: (v) => { mode = v; trackingModes.push(v); },
  });
  return { context, calls, trackingModes, table, rows, cells, tableBorders, rowBorders };
}

function setWordRunWithEnums(context) {
  global.Word = {
    run: jest.fn((cb) => cb(context)),
    RangeLocation: { start: 'Start', end: 'End', content: 'Content' },
    InsertLocation: { replace: 'Replace', after: 'After', before: 'Before' },
    ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
    Alignment: { left: 'Left', centered: 'Centered', right: 'Right', justified: 'Justified' },
    VerticalAlignment: { top: 'Top', center: 'Center', bottom: 'Bottom' },
    BorderType: { none: 'None', single: 'Single', double: 'Double' },
    BorderLocation: {
      top: 'Top', bottom: 'Bottom', left: 'Left', right: 'Right',
      insideHorizontal: 'InsideHorizontal', insideVertical: 'InsideVertical',
    },
    BuiltInStyleName: { tableGrid: 'TableGrid', gridTable4_Accent1: 'GridTable4_Accent1' },
    UnderlineType: { none: 'None', single: 'Single', double: 'Double' },
  };
}

describe('readSelectionTableRegion (style snapshot)', () => {
  beforeEach(() => jest.clearAllMocks());

  test('captures the current table style when the host exposes it', async () => {
    const context = makePrepareContext();
    const table = context.document.getSelection().parentTableOrNullObject;
    Object.assign(table, {
      styleBuiltIn: 'TableGrid',
      alignment: 'Left',
      headerRowCount: 1,
      styleBandedRows: true,
      shadingColor: null,
      font: { load: jest.fn(), bold: false, name: 'Calibri', size: 11, color: '#000000' },
      getBorder: jest.fn((loc) => ({ load: jest.fn(), type: 'Single', color: 'auto', width: 0.5 })),
    });
    setWordRun(context);

    const region = await readSelectionTableRegion(makeDeps());
    expect(region.style).toMatchObject({
      styleBuiltIn: 'TableGrid',
      alignment: 'Left',
      headerRowCount: 1,
      bandedRows: true,
      font: { bold: false, name: 'Calibri', size: 11 },
    });
    expect(region.style.borders.top).toEqual({ type: 'Single', color: 'auto', width: 0.5 });
  });

  test('degrades to a null snapshot when the style read fails', async () => {
    const context = makePrepareContext();
    const table = context.document.getSelection().parentTableOrNullObject;
    table.getBorder = jest.fn(() => { throw new Error('no borders here'); });
    setWordRun(context);

    const region = await readSelectionTableRegion(makeDeps());
    expect(region).not.toBeNull();
    expect(region.style).toBeNull();
  });
});

describe('applySelectionAmendment (style ops)', () => {
  beforeEach(() => jest.clearAllMocks());

  function styleProposal(styleOps, extra = {}) {
    return {
      selectionText: 'x',
      amendedText: null,
      commentText: null,
      tablePatch: {
        rowCount: 3,
        colCount: 2,
        cells: [],
        rowOps: [],
        merges: [],
        styleOps,
        bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
        originals: TABLE_VALUES,
        ...extra,
      },
      tableItems: [],
    };
  }

  test('three-line table: table borders + header-row bottom border', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    const result = await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'borders', borders: {
        top: { type: 'single', width: 1.5 },
        bottom: { type: 'single', width: 1.5 },
        insideH: { type: 'none' }, insideV: { type: 'none' },
      } },
      { type: 'borders', row: 1, borders: { bottom: { type: 'single', width: 0.75 } } },
    ]));

    expect(result.styleOpsApplied).toBe(2);
    expect(mock.tableBorders.Top).toEqual({ type: 'Single', color: null, width: 1.5 });
    expect(mock.tableBorders.Bottom).toEqual({ type: 'Single', color: null, width: 1.5 });
    expect(mock.tableBorders.InsideHorizontal).toEqual({ type: 'None', color: null, width: null });
    expect(mock.rowBorders['0:Bottom']).toEqual({ type: 'Single', color: null, width: 0.75 });
    expect(result.warnings).toEqual([]);
  });

  test('table style + banding flags map onto styleBuiltIn', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'tableStyle', style: 'GridTable4_Accent1', bandedRows: true, firstColumn: false },
    ]));

    expect(mock.table.styleBuiltIn).toBe('GridTable4_Accent1');
    expect(mock.table.styleBandedRows).toBe(true);
    expect(mock.table.styleFirstColumn).toBe(false);
  });

  test('header row sets headerRowCount and styles the header rows', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'headerRow', rows: 1, font: { bold: true }, shadingColor: '#DEEBF7' },
    ]));

    expect(mock.table.headerRowCount).toBe(1);
    expect(mock.rows[0].font.bold).toBe(true);
    expect(mock.rows[0].shadingColor).toBe('#DEEBF7');
    expect(mock.rows[1].font.bold).toBeNull();
  });

  test('region font runs BEFORE row structure ops (original coordinates)', async () => {
    const mock = makeStyleApplyContext();
    // Track row-2 font access order against the row delete.
    const order = [];
    const originalRow = mock.rows[1];
    Object.defineProperty(originalRow.font, 'bold', {
      get: () => null,
      set: (v) => { order.push(`fontRow2:${v}`); },
    });
    originalRow.delete.mockImplementation(() => order.push('deleteRow2'));
    setWordRunWithEnums(mock.context);

    await applySelectionAmendment(makeDeps('PC'), styleProposal(
      [{ type: 'font', region: { startRow: 2, endRow: 2, startCol: 1, endCol: 2 }, font: { bold: true } }],
      { rowOps: [{ op: 'delete', row: 2 }] },
    ));

    expect(order).toEqual(['fontRow2:true', 'deleteRow2']);
  });

  test('full-width cellFormat writes row properties; partial-width writes cells', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'cellFormat', region: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 }, shadingColor: 'red', horizontalAlignment: 'centered' },
      { type: 'cellFormat', region: { startRow: 1, endRow: 2, startCol: 2, endCol: 2 }, verticalAlignment: 'center' },
    ]));

    // Full-width → row writes (one per row), no cell shading.
    expect(mock.rows[1].shadingColor).toBe('red');
    expect(mock.rows[2].shadingColor).toBe('red');
    expect(mock.rows[1].horizontalAlignment).toBe('Centered');
    expect(mock.cells['0,0'].shadingColor).toBeNull();
    // Partial width → per-cell writes inside the band only.
    expect(mock.cells['0,1'].verticalAlignment).toBe('Center');
    expect(mock.cells['1,1'].verticalAlignment).toBe('Center');
    expect(mock.cells['2,1'].verticalAlignment).toBeNull();
    expect(mock.rows[0].verticalAlignment).toBeNull();
  });

  test('whole-table cellFormat/font map onto table properties', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'cellFormat', region: null, shadingColor: '#F3F3F3', verticalAlignment: 'center' },
      { type: 'font', region: null, font: { bold: true, size: 12 } },
    ]));

    expect(mock.table.shadingColor).toBe('#F3F3F3');
    expect(mock.table.verticalAlignment).toBe('Center');
    expect(mock.table.font.bold).toBe(true);
    expect(mock.table.font.size).toBe(12);
  });

  test('layout op drives alignment, width, autofit, and cell padding', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'layout', alignment: 'centered', widthPt: 450, autoFitWindow: true, cellPaddingPt: 5 },
    ]));

    expect(mock.table.alignment).toBe('Centered');
    expect(mock.table.width).toBe(450);
    expect(mock.table.autoFitWindow).toHaveBeenCalled();
    expect(mock.calls).toContain('padding:Top:5');
    expect(mock.calls).toContain('padding:Left:5');
    expect(mock.calls).toContain('padding:Bottom:5');
    expect(mock.calls).toContain('padding:Right:5');
  });

  test('column widths write the first-row cells', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'columnWidths', widthsPt: [120, 80] },
    ]));

    expect(mock.cells['0,0'].columnWidth).toBe(120);
    expect(mock.cells['0,1'].columnWidth).toBe(80);
  });

  test('a failing style op warns and later ops still apply', async () => {
    const mock = makeStyleApplyContext();
    mock.table.getBorder = jest.fn(() => { throw new Error('border API gone'); });
    setWordRunWithEnums(mock.context);
    const deps = makeDeps('PC');
    const result = await applySelectionAmendment(deps, styleProposal([
      { type: 'borders', borders: { all: 'single' } },
      { type: 'tableStyle', style: 'TableGrid' },
    ]));

    expect(result.styleOpsApplied).toBe(1);
    expect(result.warnings).toHaveLength(1);
    expect(mock.table.styleBuiltIn).toBe('TableGrid');
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/warning:Style op \(borders\) failed/);
  });

  test('unknown style op type is skipped with a warning', async () => {
    const mock = makeStyleApplyContext();
    setWordRunWithEnums(mock.context);
    const result = await applySelectionAmendment(makeDeps('PC'), styleProposal([
      { type: 'sparkle', magic: true },
    ]));
    expect(result.styleOpsApplied).toBe(0);
    expect(result.warnings[0]).toMatch(/Unknown style op type/);
  });
});
