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
    load: jest.fn(),
  };
  const selection = {
    parentTableOrNullObject: table,
    parentTableCellOrNullObject: { isNullObject: true, load: jest.fn() },
    getRange: jest.fn((loc) => ({
      parentTableCellOrNullObject: loc === 'Start'
        ? { isNullObject: false, rowIndex: 0, columnIndex: 0, load: jest.fn() }
        : { isNullObject: false, rowIndex: 2, columnIndex: 1, load: jest.fn() },
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
            items: [{ getRange: jest.fn(() => paraRange), load: jest.fn() }],
            load: jest.fn(),
          },
        },
        parentRow: rows[r],
      };
    }
  }

  const table = {
    isNullObject: false,
    rowCount: tableRowCount,
    load: jest.fn(),
    getCell: jest.fn((r, c) => cells[`${r},${c}`]),
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
