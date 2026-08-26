/**
 * Unit tests for the table-creation route in src/taskpane/word-actions.js
 * (prepareTableProposal / applyTableProposal, backed by lib/table-ops.js).
 *
 * Covers:
 * - Deterministic empty-table fast path for explicit dimensions (no LLM)
 * - LLM content generation with dimension hard-constraints and validation
 * - Native table insertion at document start/end and selection before/after,
 *   with platform-split change tracking
 */

jest.mock('../src/lib/llm-client.js', () => ({
  sendPrompt: jest.fn(),
  sendPromptStream: jest.fn(),
  stripMarkdown: jest.fn((t) => t),
}));

jest.mock('../src/lib/comment-extractor.js', () => ({
  extractAllComments: jest.fn(),
  extractDocumentStructured: jest.fn(),
  estimateTokenCount: jest.fn(() => 0),
  extractTrackedChanges: jest.fn(),
  extractCommentsOnRange: jest.fn(),
}));

const { sendPrompt, sendPromptStream } = require('../src/lib/llm-client.js');
const { extractDocumentStructured } = require('../src/lib/comment-extractor.js');
const {
  prepareTableProposal,
  applyTableProposal,
} = require('../src/taskpane/word-actions.js');

// --- Mock helpers ---

function makeDeps(platform = 'PC') {
  return {
    appState: {
      config: {
        backend: 'mock',
        providers: { mock: { url: '', apiKey: '', model: 'mock-model', apiPath: '' } },
        trackChangesEnabled: true,
        docExtraction: { richness: 'structured' },
      },
      platform,
    },
    log: jest.fn(),
  };
}

/** Word.run mock for prepare: the selection carries the given plain text. */
function makeSelectionContext(text = '') {
  const selection = { text, load: jest.fn() };
  return {
    document: { getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
}

/** Word.run mock for apply: records insertions and tracking-mode writes. */
function makeApplyContext() {
  const trackingModes = [];
  const inserted = [];
  const table = {
    styleBuiltIn: null,
    headerRowCount: 0,
    autoFitWindow: jest.fn(),
  };
  const body = {
    insertTable: jest.fn((r, c, loc, rows) => {
      inserted.push({ where: 'body', r, c, loc, rows });
      return table;
    }),
  };
  const selection = {
    text: '',
    load: jest.fn(),
    insertTable: jest.fn((r, c, loc, rows) => {
      inserted.push({ where: 'selection', r, c, loc, rows });
      return table;
    }),
  };
  const context = {
    document: { body, getSelection: () => selection },
    sync: jest.fn().mockResolvedValue(undefined),
  };
  let mode = null;
  Object.defineProperty(context.document, 'changeTrackingMode', {
    get: () => mode,
    set: (v) => { mode = v; trackingModes.push(v); },
  });
  return { context, table, inserted, trackingModes };
}

function setWordRun(context, { withTrackingApi = true } = {}) {
  global.Word = {
    run: jest.fn((cb) => cb(context)),
    InsertLocation: { start: 'Start', end: 'End', before: 'Before', after: 'After', replace: 'Replace' },
    ChangeTrackingMode: withTrackingApi ? { trackAll: 'TrackAll', off: 'Off' } : undefined,
    BuiltInStyleName: { tableGrid: 'TableGrid' },
  };
}

const VALID_FILLED_SPEC_JSON = JSON.stringify({
  rows: [['项目', '数量', '备注'], ['A', '1', 'x'], ['B', '2', 'y']],
  position: 'end',
  headerRowCount: 1,
  style: 'tableGrid',
  autoFit: true,
});

// --- Tests ---

describe('prepareTableProposal', () => {
  beforeEach(() => jest.clearAllMocks());

  test('explicit dimensions without content wording resolve to an empty grid, no model call', async () => {
    const proposal = await prepareTableProposal(makeDeps(), {
      instruction: '在文档的末尾插入一个三行三列的表格',
    });

    expect(sendPrompt).not.toHaveBeenCalled();
    expect(sendPromptStream).not.toHaveBeenCalled();
    expect(proposal.model).toBeNull();
    expect(proposal.spec.rows).toEqual([['', '', ''], ['', '', ''], ['', '', '']]);
    expect(proposal.spec.position).toBe('end');
    expect(proposal.warnings).toEqual([]);
  });

  test('content wording with explicit dimensions constrains the model output', async () => {
    setWordRun(makeSelectionContext('现有文档段落'));
    sendPrompt.mockResolvedValue(VALID_FILLED_SPEC_JSON);

    const proposal = await prepareTableProposal(makeDeps(), {
      instruction: '在文档末尾插入一个三行三列的表格并填充项目数据',
    });

    expect(sendPrompt).toHaveBeenCalledTimes(1);
    const promptText = sendPrompt.mock.calls[0][1];
    expect(promptText).toContain('MUST have exactly 3 rows and 3 columns');
    expect(promptText).toContain('现有文档段落');

    expect(proposal.model).toBe('mock-model');
    expect(proposal.spec.rows).toHaveLength(3);
    expect(proposal.spec.rows[0]).toEqual(['项目', '数量', '备注']);
    expect(proposal.spec.headerRowCount).toBe(1);
  });

  test('dimensionless requests let the model choose the grid', async () => {
    setWordRun(makeSelectionContext(''));
    extractDocumentStructured.mockResolvedValue('文档结构化内容');
    sendPrompt.mockResolvedValue(JSON.stringify({ rows: [['a', 'b'], ['c', 'd']] }));

    const proposal = await prepareTableProposal(makeDeps(), {
      instruction: '插入一个记录项目进度的表格',
    });

    const promptText = sendPrompt.mock.calls[0][1];
    expect(promptText).not.toContain('MUST have exactly');
    expect(promptText).toContain('文档结构化内容');
    expect(proposal.spec.rows).toEqual([['a', 'b'], ['c', 'd']]);
    expect(proposal.spec.position).toBe('end');
  });

  test('rejects model output whose dimensions contradict the instruction', async () => {
    setWordRun(makeSelectionContext(''));
    extractDocumentStructured.mockResolvedValue('');
    sendPrompt.mockResolvedValue(JSON.stringify({ rows: [['a', 'b']] }));

    await expect(prepareTableProposal(makeDeps(), {
      instruction: '插入一个三行三列的表格并填充数据',
    })).rejects.toThrow(/与要求的 3×3 不符/);
  });

  test('rejects unusable model output with actionable advice', async () => {
    setWordRun(makeSelectionContext(''));
    extractDocumentStructured.mockResolvedValue('');
    sendPrompt.mockResolvedValue('I drew you a markdown table: | a | b |');

    await expect(prepareTableProposal(makeDeps(), {
      instruction: '插入一个记录项目进度的表格',
    })).rejects.toThrow(/表格内容生成失败/);
  });

  test('streams via sendPromptStream when token handlers are given', async () => {
    setWordRun(makeSelectionContext('上下文'));
    sendPromptStream.mockResolvedValue(VALID_FILLED_SPEC_JSON);

    const proposal = await prepareTableProposal(makeDeps(), {
      instruction: '插入一个三行三列的表格并填充数据',
      onToken: jest.fn(),
    });

    expect(sendPromptStream).toHaveBeenCalled();
    expect(proposal.spec.rows).toHaveLength(3);
  });

  test('rejects an empty instruction', async () => {
    await expect(prepareTableProposal(makeDeps(), { instruction: '  ' }))
      .rejects.toThrow(/No table instruction/);
  });
});

describe('applyTableProposal', () => {
  beforeEach(() => jest.clearAllMocks());

  const emptySpec = (overrides = {}) => ({
    rows: [['', '', ''], ['', '', ''], ['', '', '']],
    position: 'end',
    headerRowCount: 0,
    style: 'tableGrid',
    autoFit: true,
    ...overrides,
  });

  test('inserts a native table at the document end, tracked on desktop', async () => {
    const { context, table, inserted, trackingModes } = makeApplyContext();
    setWordRun(context);

    const result = await applyTableProposal(makeDeps('PC'), {
      instruction: 'x',
      spec: emptySpec({ headerRowCount: 1 }),
    });

    expect(inserted).toHaveLength(1);
    expect(inserted[0]).toMatchObject({ where: 'body', r: 3, c: 3, loc: 'End' });
    expect(inserted[0].rows).toEqual(emptySpec().rows);
    expect(table.styleBuiltIn).toBe('TableGrid');
    expect(table.headerRowCount).toBe(1);
    expect(table.autoFitWindow).toHaveBeenCalled();
    expect(trackingModes).toEqual(['TrackAll', 'Off']);
    expect(result).toMatchObject({ inserted: true, rowCount: 3, columnCount: 3, tracked: true });
  });

  test('start/before/after positions pick the matching anchor and location', async () => {
    const { context, inserted } = makeApplyContext();
    setWordRun(context);

    await applyTableProposal(makeDeps('PC'), { instruction: 'x', spec: emptySpec({ position: 'start' }) });
    await applyTableProposal(makeDeps('PC'), { instruction: 'x', spec: emptySpec({ position: 'before' }) });
    await applyTableProposal(makeDeps('PC'), { instruction: 'x', spec: emptySpec({ position: 'after' }) });

    expect(inserted.map((i) => [i.where, i.loc])).toEqual([
      ['body', 'Start'],
      ['selection', 'Before'],
      ['selection', 'After'],
    ]);
  });

  test('Word for the web inserts untracked with a warning', async () => {
    const { context, trackingModes } = makeApplyContext();
    setWordRun(context);
    const deps = makeDeps('OfficeOnline');

    const result = await applyTableProposal(deps, { instruction: 'x', spec: emptySpec() });

    expect(trackingModes).toEqual(['Off', 'Off']);
    expect(result.tracked).toBe(false);
    expect(result.warnings.join('\n')).toMatch(/cannot be tracked as a revision/);
    const logged = deps.log.mock.calls.map((c) => `${c[1]}:${c[0]}`).join('\n');
    expect(logged).toMatch(/warning:Table insertion cannot be tracked/);
  });

  test('track-changes off inserts untracked without the platform warning', async () => {
    const { context, trackingModes } = makeApplyContext();
    setWordRun(context);
    const deps = makeDeps('PC');
    deps.appState.config.trackChangesEnabled = false;

    const result = await applyTableProposal(deps, { instruction: 'x', spec: emptySpec() });

    expect(trackingModes).toEqual(['Off', 'Off']);
    expect(result.tracked).toBe(false);
    expect(result.warnings).toEqual([]);
  });

  test('works without the change-tracking API (old hosts)', async () => {
    const { context, inserted } = makeApplyContext();
    setWordRun(context, { withTrackingApi: false });

    const result = await applyTableProposal(makeDeps('PC'), { instruction: 'x', spec: emptySpec() });

    expect(inserted).toHaveLength(1);
    expect(result.warnings.join('\n')).toMatch(/Change tracking is unavailable/);
  });

  test('re-validates the spec and refuses invalid proposals', async () => {
    const { context, inserted } = makeApplyContext();
    setWordRun(context);

    await expect(applyTableProposal(makeDeps('PC'), {
      instruction: 'x',
      spec: { rows: [['a'], ['b', 'c']], position: 'end' },
    })).rejects.toThrow(/表格提案无效/);
    expect(inserted).toHaveLength(0);
  });
});
