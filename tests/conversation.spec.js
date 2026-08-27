/** @jest-environment jsdom */

/**
 * Conversation turn-routing tests.
 *
 * Covers:
 *   - routeTurn: slash skill vs selection edit vs document Q&A
 *   - createConversation.submit: routes to the right word-actions handler
 *   - concurrency guard: no new turn while a run is in flight
 *   - cancel(): aborts the in-flight controllers
 *
 * word-actions is replaced by spies via deps.actions, so no Word/LLM calls
 * happen. jsdom is used because the proposal-card path touches the DOM.
 */

const { routeTurn, createConversation, TURN_TYPE, chunkCitation, looksLikeChainedInstruction } = require('../src/taskpane/conversation.js');
const { BUILTIN_SKILLS } = require('../src/taskpane/skills.js');

function makeAppState(overrides = {}) {
  return {
    isProcessing: false,
    isProcessingDoc: false,
    isProcessingSummary: false,
    processDocController: null,
    chatController: null,
    supportsComments: true,
    supportsTables: true,
    config: { commentGranularity: 0 },
    promptManager: { getPrompts: () => [], getActivePrompt: () => null },
    ...overrides,
  };
}

function makeView() {
  const msg = {
    setStatus: jest.fn(),
    setText: jest.fn(),
    appendText: jest.fn(),
    appendLogLine: jest.fn(),
    collapseLog: jest.fn(),
    appendModelToken: jest.fn(),
    collapseModelOutput: jest.fn(),
    showProgress: jest.fn(),
    hideProgress: jest.fn(),
    attachProposal: jest.fn(),
    addCitationPills: jest.fn(),
    markError: jest.fn(),
    finalizeForHistory: jest.fn(),
  };
  return {
    addUserMessage: jest.fn(),
    createAssistantMessage: jest.fn(() => msg),
    addSystemNote: jest.fn(),
    hideWelcome: jest.fn(),
    renderWelcome: jest.fn(),
    clearChat: jest.fn(),
    getCurrentSession: jest.fn(() => ({ id: 's-test', title: null, messages: [] })),
    _msg: msg,
  };
}

function makeInput() {
  return {
    setProcessing: jest.fn(),
    setValue: jest.fn(),
    focus: jest.fn(),
  };
}

function makeActions(overrides = {}) {
  return {
    hasNonEmptySelection: jest.fn(async () => false),
    readSelectionSnippet: jest.fn(async () => ''),
    prepareSelectionAmendment: jest.fn(async () => ({
      selectionText: 'before', amendedText: 'after', commentText: null, model: 'm',
    })),
    applySelectionAmendment: jest.fn(async () => {}),
    fireSelectionComment: jest.fn(async () => {}),
    runDocumentSkill: jest.fn(async () => ({
      results: [], applicationResult: { amendmentsApplied: 1, commentsInserted: 2 }, chunks: [], cancelled: false,
    })),
    runSummarySkill: jest.fn(async () => ({ chars: 10, commentCount: 0 })),
    answerQuestion: jest.fn(async () => 'the answer'),
    prepareDocumentAppend: jest.fn(async () => ({
      instruction: 'x', generatedText: 'new content', model: 'm',
    })),
    applyDocumentAppend: jest.fn(async () => ({ paragraphsAppended: 1, chars: 11 })),
    prepareEmptyParagraphCleanup: jest.fn(async () => ({ emptyCount: 3 })),
    applyEmptyParagraphCleanup: jest.fn(async () => ({ deleted: 3 })),
    prepareFormatProposal: jest.fn(async () => ({
      instruction: 'x', scope: 'selection', ops: [{ font: { bold: true } }], model: 'm',
    })),
    applyFormatProposal: jest.fn(async () => ({ applied: 1, warnings: [] })),
    prepareTableProposal: jest.fn(async () => ({
      instruction: 'x',
      spec: {
        rows: [['', '', ''], ['', '', ''], ['', '', '']],
        position: 'end', headerRowCount: 0, style: 'tableGrid', autoFit: true,
      },
      model: null,
      warnings: [],
    })),
    applyTableProposal: jest.fn(async () => ({
      inserted: true, rowCount: 3, columnCount: 3, tracked: true, warnings: [],
    })),
    prepareIllustrationProposal: jest.fn(async () => ({
      instruction: 'x',
      svg: '<svg width="10" height="10" viewBox="0 0 10 10"><rect width="10" height="10"/></svg>',
      position: 'end', model: 'm',
    })),
    prepareTableToolEdit: jest.fn(async () => null),
    prepareImageToolEdit: jest.fn(async () => ({
      instruction: 'x',
      ops: [{ type: 'delete', index: 1 }],
      items: [{ id: 1, label: 'Delete image 1', before: 'existing picture', after: '' }],
      snapshotCount: 1,
      model: 'm',
      toolLoop: { steps: 2, finished: true },
    })),
    applyImageOps: jest.fn(async () => ({ applied: 1, warnings: [] })),
    applyIllustrationProposal: jest.fn(async () => ({ inserted: true })),
    planDocumentTasks: jest.fn(async () => ({
      tasks: [{ type: 'insert', instruction: '增加标题' }, { type: 'edit', instruction: '深度润色修改' }],
      model: 'm',
    })),
    revealTextSnippet: jest.fn(async () => true),
    ...overrides,
  };
}

describe('routeTurn', () => {
  test('slash command routes to skill turn', () => {
    const turn = routeTurn('/copy-edit', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SKILL);
    expect(turn.skill.name).toBe('copy-edit');
  });

  test('free text with selection routes to selection edit', () => {
    const turn = routeTurn('make it formal', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SELECTION_EDIT);
    expect(turn.instruction).toBe('make it formal');
  });

  test('free text without selection routes to document Q&A', () => {
    const turn = routeTurn('what does clause 5 say?', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
    expect(turn.question).toBe('what does clause 5 say?');
  });

  test('edit intent without selection routes to document edit (EN)', () => {
    const turn = routeTurn('please proofread and fix this document', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_EDIT);
    expect(turn.instruction).toBe('please proofread and fix this document');
  });

  test('edit intent without selection routes to document edit (ZH)', () => {
    const turn = routeTurn('全文润色检查,请直接在word中修订', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_EDIT);
    expect(turn.instruction).toBe('全文润色检查,请直接在word中修订');
  });

  test('question lead beats edit verbs ("how should I improve...")', () => {
    const turn = routeTurn('how should I improve this section?', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('question lead beats edit verbs (ZH "如何修改")', () => {
    const turn = routeTurn('如何修改第三章更好?', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('edit intent with a selection still routes to selection edit', () => {
    const turn = routeTurn('polish this', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('question with a selection routes to Q&A (selection becomes context)', () => {
    const turn = routeTurn('如何评价这段话?', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
    expect(turn.question).toBe('如何评价这段话?');
  });

  test('question lead with selection routes to Q&A (EN)', () => {
    const turn = routeTurn('what does this clause mean?', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('review intent with selection routes to Q&A, not the edit pipeline', () => {
    // The reported scenario: "check the selected table's contents" is an
    // analysis request — routing it to SELECTION_EDIT made it hit the
    // (former) uniform-table guard and error out.
    const turn = routeTurn('检查选择的表格的内容', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
    expect(turn.question).toBe('检查选择的表格的内容');

    expect(routeTurn('check the selected data', { hasSelection: true, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOC_QA);
  });

  test('review + edit verbs with selection still route to selection edit', () => {
    const turn = routeTurn('检查并修改这段话', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('image-only selection routes every instruction to the image tool session', () => {
    // Edit-ish instruction: the text pipelines have nothing to operate on —
    // the image OBJECT (snapshot index + tools) is the context.
    expect(routeTurn('润色一下', { hasSelection: true, hasImageSelection: true, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    // Questions too: visual reading is the read_image tool, not injected bytes.
    expect(routeTurn('这是什么图?', { hasSelection: true, hasImageSelection: true, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('describe this picture', { hasSelection: true, hasImageSelection: true, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
  });

  test('mixed text+image selection keeps text routing (edit/question)', () => {
    expect(routeTurn('make it formal', { hasSelection: true, hasImageSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.SELECTION_EDIT);
    expect(routeTurn('what does this say?', { hasSelection: true, hasImageSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOC_QA);
  });

  test('image-only selection takes document scope for format intent', () => {
    const turn = routeTurn('把标题居中', { hasSelection: true, hasImageSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.FORMAT);
    expect(turn.scope).toBe('document');
    // Text selection keeps selection scope.
    expect(routeTurn('把标题居中', { hasSelection: true, skills: BUILTIN_SKILLS }).scope).toBe('selection');
  });

  test('empty input returns null', () => {
    expect(routeTurn('   ', { hasSelection: false, skills: BUILTIN_SKILLS })).toBeNull();
    expect(routeTurn('', { hasSelection: true, skills: BUILTIN_SKILLS })).toBeNull();
  });

  test('append intent without selection routes to doc-append (ZH)', () => {
    const turn = routeTurn('继续丰富内容，并追加到文档末尾', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_APPEND);
    expect(turn.instruction).toBe('继续丰富内容，并追加到文档末尾');
  });

  test('append intent without selection routes to doc-append (EN)', () => {
    const turn = routeTurn('append a closing paragraph to the document', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_APPEND);
  });

  test('append intent wins over a selection', () => {
    const turn = routeTurn('续写这段并追加到文档末尾', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_APPEND);
  });

  test('question lead beats append intent (ZH "如何续写")', () => {
    const turn = routeTurn('如何续写这篇文章?', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('edit intent without append keywords still routes to document edit', () => {
    const turn = routeTurn('润色全文', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_EDIT);
  });

  test('format intent with a selection routes to format (selection scope, ZH)', () => {
    const turn = routeTurn('把这段话加粗并标红', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.FORMAT);
    expect(turn.instruction).toBe('把这段话加粗并标红');
    expect(turn.scope).toBe('selection');
  });

  test('format intent without a selection routes to format (document scope, ZH)', () => {
    const turn = routeTurn('把所有标题设为居中', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.FORMAT);
    expect(turn.scope).toBe('document');
  });

  test('"add a title" routes to format — insert ops live in the format pipeline', () => {
    const turn = routeTurn('增加文章标题', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.FORMAT);
    expect(turn.scope).toBe('document');
  });

  test('"add title + restyle document" hits format+edit families -> compound turn', () => {
    const turn = routeTurn('增加标题，全文样式设计修改', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
    expect(turn.instruction).toBe('增加标题，全文样式设计修改');
  });

  test('"add title + deep polish" routes to compound (ZH)', () => {
    const turn = routeTurn('增加标题，并深度润色修改', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
  });

  test('"continue writing + center everything" routes to compound (append+format)', () => {
    const turn = routeTurn('续写第三章并把全文居中', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
  });

  test('compound detection also applies with a selection (format+edit verbs)', () => {
    const turn = routeTurn('把这段话加粗并润色', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
  });

  test('single-intent instructions never route to compound', () => {
    expect(routeTurn('增加文章标题', { hasSelection: false, skills: BUILTIN_SKILLS }).type).toBe(TURN_TYPE.FORMAT);
    expect(routeTurn('润色这段话', { hasSelection: true, skills: BUILTIN_SKILLS }).type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('question lead beats compound detection', () => {
    const turn = routeTurn('如何润色并加标题？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('"enrich content and update the document" hits edit intent (ZH)', () => {
    const turn = routeTurn('丰富内容并更新文档', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_EDIT);
    expect(turn.instruction).toBe('丰富内容并更新文档');
  });

  test('update/enrich verbs need a document-ish object ("是谁更新的" stays out)', () => {
    // No edit intent hit, no question lead: ambiguous -> planner classify.
    const turn = routeTurn('文档最后是谁更新的', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
  });

  test('ambiguous zero-hit instruction routes to the planner for classification', () => {
    const turn = routeTurn('让文章更有感染力', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
    expect(turn.instruction).toBe('让文章更有感染力');
  });

  test('ambiguous zero-hit instruction falls back to Q&A when compound is disabled', () => {
    const turn = routeTurn('让文章更有感染力', { hasSelection: false, skills: BUILTIN_SKILLS, allowCompound: false });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
    expect(turn.question).toBe('让文章更有感染力');
  });

  test('question lead still skips the classifier (no planner call)', () => {
    const turn = routeTurn('如何评价这篇文章？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('illustration intent routes to illustration (ZH)', () => {
    const turn = routeTurn('设计并增加SVG插图', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.ILLUSTRATION);
    expect(turn.instruction).toBe('设计并增加SVG插图');
  });

  test('illustration intent routes to illustration (EN)', () => {
    const turn = routeTurn('add an image of a cat', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.ILLUSTRATION);
  });

  test('illustration intent wins over a selection', () => {
    const turn = routeTurn('给这篇文章配一张插图', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.ILLUSTRATION);
  });

  test('image management routes to the image tool loop (ZH + EN)', () => {
    expect(routeTurn('删除文档里的第二张图片', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('把第一张图片替换成一张黄昏场景', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('delete the second image', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('resize image 1 to 300pt', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
  });

  test('multi-image design routes to the image tool loop', () => {
    expect(routeTurn('设计两张插图，一张在开头一张在文末', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    // Single-image design stays on the dedicated illustration turn.
    expect(routeTurn('设计一张插图放在文末', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.ILLUSTRATION);
  });

  test('image questions stay Q&A even with management verbs', () => {
    expect(routeTurn('如何删除文档里的图片？', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOC_QA);
  });

  test('chained instruction detection', () => {
    expect(looksLikeChainedInstruction('删除空行，然后重新编号')).toBe(true);
    expect(looksLikeChainedInstruction('先改标题，接着润色全文')).toBe(true);
    expect(looksLikeChainedInstruction('加一行，再删一行')).toBe(true);
    expect(looksLikeChainedInstruction('再润色一下')).toBe(false);
    expect(looksLikeChainedInstruction('润色这段话')).toBe(false);
  });

  test('generic image words without a creation verb stay an edit', () => {
    const turn = routeTurn('修改图像描述的措辞', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('question lead beats illustration intent ("如何给文章配插图？")', () => {
    const turn = routeTurn('如何给文章配插图？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('format intent routes to format (EN)', () => {
    const turn = routeTurn('make this bold and set it to Heading 2', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.FORMAT);
    expect(turn.scope).toBe('selection');
  });

  test('question lead beats format intent ("如何修改样式？")', () => {
    const turn = routeTurn('如何修改样式？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('non-format instruction with a selection still routes to selection edit', () => {
    const turn = routeTurn('润色这段话', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('empty-paragraph cleanup routes to cleanup (ZH)', () => {
    const turn = routeTurn('删除多余的空段落', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.CLEANUP);
    expect(turn.instruction).toBe('删除多余的空段落');
  });

  test('empty-paragraph cleanup routes to cleanup (EN)', () => {
    const turn = routeTurn('delete all empty paragraphs', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.CLEANUP);
  });

  test('cleanup intent is document scope even with a selection', () => {
    const turn = routeTurn('清除文档里多余的空行', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.CLEANUP);
  });

  test('question lead beats cleanup intent (ZH "为什么有多余的空段落？")', () => {
    const turn = routeTurn('为什么有多余的空段落？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('cleanup + edit hits two families -> compound', () => {
    const turn = routeTurn('删除多余的空段落，然后润色全文', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
  });

  test('table creation intent routes to table (ZH, explicit dimensions)', () => {
    const turn = routeTurn('在文档的末尾插入一个三行三列的表格', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.TABLE);
    expect(turn.instruction).toBe('在文档的末尾插入一个三行三列的表格');
  });

  test('table placement phrasing is not counted as an append request', () => {
    // "到文档末尾" names WHERE the table goes — a single table intent, not compound.
    const turn = routeTurn('到文档末尾插入一个表格', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.TABLE);
  });

  test('table intent routes to table (EN) and beats append keywords', () => {
    const turn = routeTurn('add a 3x3 table to the document', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.TABLE);
  });

  test('table intent wins over a selection (anchors before/after it)', () => {
    const turn = routeTurn('在选中的段落后面插入一个表格', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.TABLE);
  });

  test('editing mention of an existing table stays off the table route', () => {
    const turn = routeTurn('修改表格第二行的内容', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('question lead beats table intent ("如何在文档中插入表格？")', () => {
    const turn = routeTurn('如何在文档中插入表格？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
  });

  test('table + polish hits two families -> compound', () => {
    const turn = routeTurn('插入一个三行三列的表格，并深度润色修改', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
  });
});

describe('createConversation.submit', () => {
  test('free text + selection runs the selection-edit pipeline (proposal card, no direct apply)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const input = makeInput();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input, log: jest.fn(),
      actions, getSelectionText: async () => 'some selected text',
    });

    await conv.submit('make it formal');

    expect(actions.prepareSelectionAmendment).toHaveBeenCalledTimes(1);
    expect(actions.prepareSelectionAmendment.mock.calls[0][1].promptTemplate).toBe('make it formal');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    // Never applied directly — the proposal card gates the document write.
    expect(actions.applySelectionAmendment).not.toHaveBeenCalled();
    expect(input.setProcessing).toHaveBeenCalledWith(true);
    expect(input.setProcessing).toHaveBeenLastCalledWith(false);
  });

  test('chained instruction + selection tries the table tool loop first', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareTableToolEdit: jest.fn(async () => ({
        selectionText: 'a\nb', amendedText: null, commentText: null, model: 'm',
        tablePatch: {
          rowCount: 2, colCount: 1,
          cells: [{ row: 2, col: 1, text: 'new' }],
          rowOps: [],
          bounds: { startRow: 1, endRow: 2, startCol: 1, endCol: 1 },
          originals: [['a'], ['b']],
        },
        tableItems: [{ label: 'Cell R2C1', before: 'b', after: 'new' }],
        toolLoop: { steps: 3, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'a b',
    });

    await conv.submit('把第二格改好，然后加一行合计');

    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(actions.prepareTableToolEdit.mock.calls[0][1].instruction).toBe('把第二格改好，然后加一行合计');
    // Single-shot path untouched when the tool loop produced the proposal.
    expect(actions.prepareSelectionAmendment).not.toHaveBeenCalled();
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
  });

  test('chained instruction on a non-table selection falls back to single-shot', async () => {
    const actions = makeActions({ prepareTableToolEdit: jest.fn(async () => null) });
    const conv = createConversation({
      appState: makeAppState(), view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'plain text',
    });

    await conv.submit('先润色，然后压缩这段话');

    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(actions.prepareSelectionAmendment).toHaveBeenCalledTimes(1);
  });

  test('unparseable single-shot table patch retries via the tool loop', async () => {
    const toolProposal = {
      selectionText: 'a\nb', amendedText: null, commentText: null, model: 'm',
      tablePatch: {
        rowCount: 2, colCount: 1,
        cells: [{ row: 1, col: 1, text: 'fixed' }],
        rowOps: [], bounds: { startRow: 1, endRow: 2, startCol: 1, endCol: 1 },
        originals: [['a'], ['b']],
      },
      tableItems: [{ label: 'Cell R1C1', before: 'a', after: 'fixed' }],
      toolLoop: { steps: 4, finished: true },
    };
    const actions = makeActions({
      prepareSelectionAmendment: jest.fn(async () => {
        throw new Error('Table patch response contains no JSON object');
      }),
      prepareTableToolEdit: jest.fn(async () => toolProposal),
    });
    const view = makeView();
    const conv = createConversation({
      appState: makeAppState(), view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'a b',
    });

    await conv.submit('修改表格内容');

    expect(actions.prepareSelectionAmendment).toHaveBeenCalledTimes(1);
    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    expect(view._msg.markError).not.toHaveBeenCalled();
  });

  test('image tool turn stages ops and apply runs only checked ops', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('删除文档里的第二张图片');

    expect(actions.prepareImageToolEdit).toHaveBeenCalledTimes(1);
    expect(actions.prepareImageToolEdit.mock.calls[0][1].instruction).toBe('删除文档里的第二张图片');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);

    // One checkbox per op; checking it and applying runs exactly that op.
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    const boxes = cardEl.querySelectorAll('.proposal-card-change input[type="checkbox"]');
    expect(boxes).toHaveLength(1);
    boxes[0].checked = true;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyImageOps).toHaveBeenCalledTimes(1);
    expect(actions.applyImageOps.mock.calls[0][1].ops).toEqual([{ type: 'delete', index: 1 }]);
  });

  test('image-only selection enters the image tool session with selection focus', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions,
      getSelectionContent: async () => ({
        text: '',
        totalImages: 1,
        images: [{ base64: 'HUGE', dataUrl: 'data:image/png;base64,HUGE', width: 300, height: 200, altText: 'sunset' }],
      }),
    });

    await conv.submit('描述这张图');

    expect(actions.prepareImageToolEdit).toHaveBeenCalledTimes(1);
    const args = actions.prepareImageToolEdit.mock.calls[0][1];
    expect(args.instruction).toBe('描述这张图');
    // Metadata only — the base64 payload never leaves the selection reader.
    expect(args.selectionImages).toEqual([{ width: 300, height: 200, altText: 'sunset' }]);
  });

  test('image tool loop with a read-only outcome answers in chat (no card)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareImageToolEdit: jest.fn(async () => ({
        noOps: true,
        answer: 'the picture shows a bar chart of revenue',
        ops: [],
        items: [],
        snapshotCount: 1,
        model: 'm',
        toolLoop: { steps: 2, finished: true },
      })),
    });
    // Image-only selection: any instruction enters the image tool session.
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions,
      getSelectionContent: async () => ({
        text: '',
        totalImages: 1,
        images: [{ dataUrl: 'data:image/png;base64,X', width: 320, height: 200, altText: 'chart' }],
      }),
    });

    await conv.submit('描述这张图');

    expect(actions.prepareImageToolEdit).toHaveBeenCalledTimes(1);
    expect(view._msg.setText).toHaveBeenCalledWith('the picture shows a bar chart of revenue');
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
  });

  test('mixed text+image question carries image metadata into Q&A', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions,
      getSelectionContent: async () => ({
        text: 'Figure 1 shows revenue.',
        totalImages: 1,
        images: [{ dataUrl: 'data:image/png;base64,X', width: 300, height: 200, altText: 'revenue chart' }],
      }),
    });

    // "解释" leads with a question marker — routes mixed selection to DOC_QA.
    await conv.submit('解释这段和图');

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    const args = actions.answerQuestion.mock.calls[0][1];
    expect(args.selectionText).toBe('Figure 1 shows revenue.');
    expect(args.selectionImages).toEqual([{ width: 300, height: 200, altText: 'revenue chart' }]);
  });

  test('free text without selection runs document Q&A', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('what is this document about?');

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    expect(actions.answerQuestion.mock.calls[0][1].question).toBe('what is this document about?');
    expect(view._msg.setText).toHaveBeenCalledWith('the answer');
  });

  test('append intent stages an append proposal (no direct write)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('继续丰富内容，并追加到文档末尾');

    expect(actions.prepareDocumentAppend).toHaveBeenCalledTimes(1);
    expect(actions.prepareDocumentAppend.mock.calls[0][1].instruction).toBe('继续丰富内容，并追加到文档末尾');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    // Staged: nothing written to the document before the user clicks Apply.
    expect(actions.applyDocumentAppend).not.toHaveBeenCalled();
    expect(actions.answerQuestion).not.toHaveBeenCalled();

    // Clicking "Apply as tracked changes" writes the generated content.
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyDocumentAppend).toHaveBeenCalledTimes(1);
  });

  test('append turn with empty model output shows a status instead of a card', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareDocumentAppend: jest.fn(async () => ({ instruction: 'x', generatedText: '', model: 'm' })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('追加一段结尾到文档末尾');

    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.setStatus).toHaveBeenCalledWith('The model returned no content to append.');
    expect(actions.applyDocumentAppend).not.toHaveBeenCalled();
  });

  test('table intent stages a table proposal with a grid preview (no direct insert)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('在文档的末尾插入一个三行三列的表格');

    expect(actions.prepareTableProposal).toHaveBeenCalledTimes(1);
    expect(actions.prepareTableProposal.mock.calls[0][1].instruction).toBe('在文档的末尾插入一个三行三列的表格');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    // Staged: the table preview is attached to the persisted meta, and nothing
    // is inserted before the user clicks Apply.
    const meta = view._msg.attachProposal.mock.calls[0][1];
    expect(meta.tablePreview.rows).toHaveLength(3);
    expect(meta.tablePreview.rows[0]).toHaveLength(3);
    expect(meta.tablePreview.position).toBe('end');
    expect(actions.applyTableProposal).not.toHaveBeenCalled();
    expect(actions.answerQuestion).not.toHaveBeenCalled();

    // Clicking "Apply as tracked changes" inserts the table.
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyTableProposal).toHaveBeenCalledTimes(1);
  });

  test('table apply warnings settle the card into a warning state', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      applyTableProposal: jest.fn(async () => ({
        inserted: true, rowCount: 3, columnCount: 3, tracked: false,
        warnings: ['Table insertion cannot be tracked as a revision on this host (OfficeOnline) — applied directly.'],
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('add a 3x3 table to the document');
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));

    expect(actions.applyTableProposal).toHaveBeenCalledTimes(1);
    expect(cardEl.className).toMatch(/proposal-warning/);
  });

  test('table turn reports the requirement on hosts without WordApi 1.3', async () => {
    const appState = makeAppState({ supportsTables: false });
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('在文档的末尾插入一个三行三列的表格');

    expect(actions.prepareTableProposal).not.toHaveBeenCalled();
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.markError).toHaveBeenCalledTimes(1);
    expect(view._msg.markError.mock.calls[0][0]).toMatch(/WordApi 1\.3/);
  });

  test('table turn surfaces prepare failures as message errors', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareTableProposal: jest.fn(async () => { throw new Error('表格内容生成失败（…）'); }),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('插入一个记录项目进度的表格');

    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.markError).toHaveBeenCalledWith('表格内容生成失败（…）');
  });

  test('cleanup intent stages an empty-paragraph deletion proposal (no direct delete)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('删除多余的空段落');

    expect(actions.prepareEmptyParagraphCleanup).toHaveBeenCalledTimes(1);
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    // Staged: nothing deleted before the user clicks Apply.
    expect(actions.applyEmptyParagraphCleanup).not.toHaveBeenCalled();
    expect(actions.answerQuestion).not.toHaveBeenCalled();

    // Clicking "Apply as tracked changes" deletes the empty paragraphs.
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyEmptyParagraphCleanup).toHaveBeenCalledTimes(1);
  });

  test('cleanup turn with no empty paragraphs shows a status instead of a card', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareEmptyParagraphCleanup: jest.fn(async () => ({ emptyCount: 0 })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('删除多余的空段落');

    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.setStatus).toHaveBeenCalledWith('No empty paragraphs found.');
    expect(actions.applyEmptyParagraphCleanup).not.toHaveBeenCalled();
  });

  test('cleanup apply that deleted nothing settles into a warning state', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareEmptyParagraphCleanup: jest.fn(async () => ({ emptyCount: 2 })),
      applyEmptyParagraphCleanup: jest.fn(async () => ({ deleted: 0 })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('删除多余的空段落');

    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(cardEl.classList.contains('proposal-warning')).toBe(true);
  });

  test('planned edit task that is really an empty-paragraph cleanup routes to the cleanup pipeline', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({
        tasks: [
          { type: 'edit', instruction: '删除文档中多余的空段落' },
          { type: 'edit', instruction: '深度润色修改' },
        ],
        model: 'm',
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('删除多余的空段落，然后润色全文');

    expect(actions.prepareEmptyParagraphCleanup).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill).toHaveBeenCalledTimes(1);
    expect(actions.answerQuestion).not.toHaveBeenCalled();
  });

  test('question with a selection runs Q&A with the selection as context', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '  the selected clause  ',
    });

    await conv.submit('what does this clause mean?');

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    expect(actions.answerQuestion.mock.calls[0][1].question).toBe('what does this clause mean?');
    expect(actions.answerQuestion.mock.calls[0][1].selectionText).toBe('the selected clause');
    // Must not run the selection-edit pipeline for a question
    expect(actions.prepareSelectionAmendment).not.toHaveBeenCalled();
  });

  test('edit intent without selection stages a document proposal (no direct apply)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [{ status: 'fulfilled', amendment: 'new text', chunk: { id: 'c0', text: 'old text' } }],
      chunks: [{ id: 'c0', paragraphs: [{ text: 'old text' }] }],
      apply: jest.fn(async () => ({ amendmentsApplied: 1, commentsInserted: 0 })),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({
      runDocumentSkill: jest.fn(async () => staged),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');

    expect(actions.runDocumentSkill).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill.mock.calls[0][1].category).toBe('amendment');
    expect(actions.runDocumentSkill.mock.calls[0][1].promptTemplate).toBe('please polish the whole document');
    expect(actions.runDocumentSkill.mock.calls[0][1].gateApply).toBe(true);
    // Staged: proposal card shown, nothing applied yet.
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    expect(staged.apply).not.toHaveBeenCalled();
    expect(actions.answerQuestion).not.toHaveBeenCalled();

    // Clicking "Apply as tracked changes" applies the staged run.
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(staged.apply).toHaveBeenCalledTimes(1);
    expect(view._msg.addCitationPills).toHaveBeenCalledTimes(1);
  });

  test('staged document proposal with no amendments is discarded', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [{ status: 'fulfilled', amendment: null, chunk: { id: 'c0', text: 'same' } }],
      chunks: [{ id: 'c0', paragraphs: [{ text: 'same' }] }],
      apply: jest.fn(),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({
      runDocumentSkill: jest.fn(async () => staged),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');

    expect(staged.discard).toHaveBeenCalledTimes(1);
    expect(staged.apply).not.toHaveBeenCalled();
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.setStatus).toHaveBeenCalledWith('The model proposed no changes.');
  });

  test('amendment identical to the original chunk text is discarded (LLM echo)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [{
        status: 'fulfilled',
        amendment: 'Para one\nPara two',
        chunk: { id: 'c0', paragraphs: [{ text: 'Para one' }, { text: 'Para two' }] },
      }],
      chunks: [{ id: 'c0', paragraphs: [{ text: 'Para one' }, { text: 'Para two' }] }],
      apply: jest.fn(),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({
      runDocumentSkill: jest.fn(async () => staged),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');

    expect(staged.discard).toHaveBeenCalledTimes(1);
    expect(staged.apply).not.toHaveBeenCalled();
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.setStatus).toHaveBeenCalledWith('The model proposed no changes.');
  });

  test('summary skill routes to the summary pipeline', async () => {
    const appState = makeAppState();
    const actions = makeActions();
    const conv = createConversation({
      appState, view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/summarize-contract');

    expect(actions.runSummarySkill).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
  });

  test('document-scope skill runs the document pipeline with citation pills', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      runDocumentSkill: jest.fn(async () => ({
        results: [],
        applicationResult: { amendmentsApplied: 0, commentsInserted: 3 },
        chunks: [{ id: 'c0', paragraphs: [{ text: 'Section 1. Definitions and Interpretation' }] }],
        cancelled: false,
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/check-doc');

    expect(actions.runDocumentSkill).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill.mock.calls[0][1].category).toBe('comment');
    expect(view._msg.addCitationPills).toHaveBeenCalledTimes(1);
    const pills = view._msg.addCitationPills.mock.calls[0][0];
    expect(pills[0].label).toBe('Section 1. Definitions and Interpretation');
    expect(appState.isProcessingDoc).toBe(false);
  });

  test('selection-first amendment skill with a selection stages a proposal', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'some selected text',
    });

    await conv.submit('/copy-edit');

    expect(actions.prepareSelectionAmendment).toHaveBeenCalledTimes(1);
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
  });

  test('selection-first amendment skill without a selection falls back to document scope', async () => {
    const appState = makeAppState();
    const actions = makeActions();
    const conv = createConversation({
      appState, view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/copy-edit');

    expect(actions.runDocumentSkill).toHaveBeenCalledTimes(1);
    expect(actions.prepareSelectionAmendment).not.toHaveBeenCalled();
  });

  test('chat skill answers in chat via the Q&A path', async () => {
    const appState = makeAppState();
    const actions = makeActions();
    const conv = createConversation({
      appState, view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/industry-overview tell me about SaaS');

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    expect(actions.answerQuestion.mock.calls[0][1].question).toBe('tell me about SaaS');
    expect(actions.answerQuestion.mock.calls[0][1].skillTemplate).toContain('industry analyst');
  });

  test('concurrency guard: submit while a doc run is in flight is rejected', async () => {
    const appState = makeAppState({ isProcessingDoc: true });
    const log = jest.fn();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log,
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/check-doc');

    expect(log).toHaveBeenCalledWith(expect.stringContaining('Already processing'), 'warning');
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
    expect(view.addUserMessage).not.toHaveBeenCalled();
  });

  test('handler errors surface on the message instead of throwing', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      runSummarySkill: jest.fn(async () => { throw new Error('boom'); }),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/summarize-contract');

    expect(view._msg.markError).toHaveBeenCalledWith(expect.stringContaining('boom'));
    expect(appState.isProcessingSummary).toBe(false);
  });

  test('format intent stages a format proposal (no direct write, apply on click)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'some selected text',
    });

    await conv.submit('把这段话加粗并标红');

    expect(actions.prepareFormatProposal).toHaveBeenCalledTimes(1);
    expect(actions.prepareFormatProposal.mock.calls[0][1].instruction).toBe('把这段话加粗并标红');
    expect(actions.prepareFormatProposal.mock.calls[0][1].scope).toBe('selection');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    // Staged: nothing written to the document before the user clicks Apply.
    expect(actions.applyFormatProposal).not.toHaveBeenCalled();
    expect(actions.prepareSelectionAmendment).not.toHaveBeenCalled();
    expect(actions.answerQuestion).not.toHaveBeenCalled();

    // Clicking "Apply as tracked changes" applies the formatting ops.
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyFormatProposal).toHaveBeenCalledTimes(1);
  });

  test('format turn with no parsed ops shows a status instead of a card', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareFormatProposal: jest.fn(async () => ({ instruction: 'x', scope: 'document', ops: [], model: 'm' })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('把所有标题设为居中');

    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.setStatus).toHaveBeenCalledWith('The model proposed no changes.');
    expect(actions.applyFormatProposal).not.toHaveBeenCalled();
  });

  test('illustration turn stages a previewable proposal (apply on click)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('设计并增加SVG插图');

    expect(actions.prepareIllustrationProposal).toHaveBeenCalledTimes(1);
    expect(actions.prepareIllustrationProposal.mock.calls[0][1].instruction).toBe('设计并增加SVG插图');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    // Staged: nothing written to the document before the user clicks Apply.
    expect(actions.applyIllustrationProposal).not.toHaveBeenCalled();
    expect(actions.prepareDocumentAppend).not.toHaveBeenCalled();
    expect(actions.answerQuestion).not.toHaveBeenCalled();

    // The staged card carries an image preview of the proposed artwork.
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    expect(cardEl.querySelector('img.proposal-card-preview')).not.toBeNull();

    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyIllustrationProposal).toHaveBeenCalledTimes(1);
  });

  test('illustration turn with no usable SVG shows a status instead of a card', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareIllustrationProposal: jest.fn(async () => ({ instruction: 'x', svg: null, position: 'end', model: 'm' })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('设计并增加SVG插图');

    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.setStatus).toHaveBeenCalledWith('The model produced no usable SVG illustration.');
    expect(actions.applyIllustrationProposal).not.toHaveBeenCalled();
  });

  test('compound turn plans tasks and stages one proposal per pipeline', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [{ status: 'fulfilled', amendment: 'new text', chunk: { id: 'c0', text: 'old text' } }],
      chunks: [{ id: 'c0', paragraphs: [{ text: 'old text' }] }],
      apply: jest.fn(async () => ({ amendmentsApplied: 1, commentsInserted: 0 })),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({ runDocumentSkill: jest.fn(async () => staged) });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('增加标题，并深度润色修改');

    expect(actions.planDocumentTasks).toHaveBeenCalledTimes(1);
    expect(actions.planDocumentTasks.mock.calls[0][1].instruction).toBe('增加标题，并深度润色修改');
    // insert task -> format pipeline at document scope (a title never
    // inserts into a selection); edit task -> whole-document amendment run.
    expect(actions.prepareFormatProposal).toHaveBeenCalledTimes(1);
    expect(actions.prepareFormatProposal.mock.calls[0][1].instruction).toBe('增加标题');
    expect(actions.prepareFormatProposal.mock.calls[0][1].scope).toBe('document');
    expect(actions.runDocumentSkill).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill.mock.calls[0][1].promptTemplate).toBe('深度润色修改');
    // One card per task; nothing applied before the user clicks Apply.
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(2);
    expect(actions.applyFormatProposal).not.toHaveBeenCalled();
    expect(staged.apply).not.toHaveBeenCalled();
  });

  test('compound turn with a selection scopes edit/format tasks to the selection', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({
        tasks: [{ type: 'format', instruction: '加粗' }, { type: 'edit', instruction: '润色' }],
        model: 'm',
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'some selected text',
    });

    await conv.submit('把这段话加粗并润色');

    expect(actions.prepareFormatProposal).toHaveBeenCalledTimes(1);
    expect(actions.prepareFormatProposal.mock.calls[0][1].scope).toBe('selection');
    expect(actions.prepareSelectionAmendment).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
  });

  test('compound turn falls back to single-intent routing when planning fails', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({ tasks: null, model: 'm' })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('增加标题，并深度润色修改');

    // Single-intent fallback: 标题 -> format pipeline (document scope).
    expect(actions.prepareFormatProposal).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
  });

  test('ambiguous zero-hit instruction goes through the planner, then the classified pipeline runs', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({
        tasks: [{ type: 'edit', instruction: '丰富内容并更新文档' }],
        model: 'm',
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('让文章更有感染力');

    expect(actions.planDocumentTasks).toHaveBeenCalledTimes(1);
    expect(actions.planDocumentTasks.mock.calls[0][1].instruction).toBe('让文章更有感染力');
    // The planner classified it as an edit -> document amendment pipeline.
    expect(actions.runDocumentSkill).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill.mock.calls[0][1].promptTemplate).toBe('丰富内容并更新文档');
    expect(actions.answerQuestion).not.toHaveBeenCalled();
  });

  test('ambiguous input falls back to Q&A when planning fails', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({ tasks: null, model: 'm' })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('让文章更有感染力');

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    expect(actions.answerQuestion.mock.calls[0][1].question).toBe('让文章更有感染力');
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
  });

  test('format card lists one selectable change per op; apply runs only checked ops', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareFormatProposal: jest.fn(async () => ({
        instruction: 'x', scope: 'selection',
        ops: [
          { match: 'a', font: { bold: true } },
          { paragraphStyle: 'heading1', paragraph: { alignment: 'centered' } },
        ],
        model: 'm',
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'some selected text',
    });

    await conv.submit('把这段话加粗并标红');

    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    const boxes = cardEl.querySelectorAll('.proposal-card-change input[type="checkbox"]');
    expect(boxes).toHaveLength(2);
    expect(cardEl.textContent).toContain('"a" → bold');
    expect(cardEl.textContent).toContain('heading1 paragraphs → alignment: centered');

    // Uncheck the second op; Apply must run only the first.
    boxes[1].checked = false;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyFormatProposal).toHaveBeenCalledTimes(1);
    expect(actions.applyFormatProposal.mock.calls[0][1].ops)
      .toEqual([{ match: 'a', font: { bold: true } }]);
    expect(cardEl.querySelector('.proposal-card-status').textContent).toContain('1 of 2');
  });

  test('document edit card lists one selectable change per section', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [
        { status: 'fulfilled', amendment: 'new zero', chunk: { id: 'c0', paragraphs: [{ text: 'old zero' }] } },
        { status: 'fulfilled', amendment: 'new one', chunk: { id: 'c1', paragraphs: [{ text: 'old one' }] } },
      ],
      chunks: [
        { id: 'c0', paragraphs: [{ text: 'old zero' }] },
        { id: 'c1', paragraphs: [{ text: 'old one' }] },
      ],
      apply: jest.fn(async () => ({ amendmentsApplied: 1, commentsInserted: 0 })),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({ runDocumentSkill: jest.fn(async () => staged) });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');

    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    const boxes = cardEl.querySelectorAll('.proposal-card-change input[type="checkbox"]');
    expect(boxes).toHaveLength(2);
    // Each section shows an inline before/after diff (word-level).
    const diffs = cardEl.querySelectorAll('.diff-view');
    expect(diffs).toHaveLength(2);
    expect(diffs[0].querySelector('del').textContent).toBe('old');
    expect(diffs[0].querySelector('ins').textContent).toBe('new');

    // Uncheck section c1; Apply must carry only c0's chunk id.
    boxes[1].checked = false;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(staged.apply).toHaveBeenCalledWith(['c0']);
  });

  test('document edit card warns honestly when apply lands nothing (skipped chunk)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [{ status: 'fulfilled', amendment: 'new text', chunk: { id: 'c0', text: 'old text' } }],
      chunks: [{ id: 'c0', paragraphs: [{ text: 'old text' }] }],
      apply: jest.fn(async () => ({
        amendmentsApplied: 0,
        commentsInserted: 0,
        errors: ['Chunk c0: original content no longer matches the staged range (edited since staging?); amendment skipped'],
      })),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({ runDocumentSkill: jest.fn(async () => staged) });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));

    expect(staged.apply).toHaveBeenCalledTimes(1);
    // The card must NOT claim success when nothing landed.
    expect(cardEl.classList.contains('proposal-applied')).toBe(false);
    expect(cardEl.classList.contains('proposal-warning')).toBe(true);
    const statusText = cardEl.querySelector('.proposal-card-status').textContent;
    expect(statusText).toContain('Nothing applied');
    expect(statusText).toContain('no longer matches');
    expect(view._msg.addCitationPills).not.toHaveBeenCalled();
  });

  test('document edit card warns when the staged edits already match the document', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [{ status: 'fulfilled', amendment: 'new text', chunk: { id: 'c0', text: 'old text' } }],
      chunks: [{ id: 'c0', paragraphs: [{ text: 'old text' }] }],
      apply: jest.fn(async () => ({ amendmentsApplied: 0, commentsInserted: 0, errors: [] })),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({ runDocumentSkill: jest.fn(async () => staged) });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));

    expect(cardEl.classList.contains('proposal-warning')).toBe(true);
    expect(cardEl.querySelector('.proposal-card-status').textContent)
      .toContain('already match the document');
  });

  test('document edit card reports skipped sections alongside successful applies', async () => {
    const appState = makeAppState();
    const view = makeView();
    const staged = {
      staged: true,
      results: [{ status: 'fulfilled', amendment: 'new text', chunk: { id: 'c0', text: 'old text' } }],
      chunks: [{ id: 'c0', paragraphs: [{ text: 'old text' }] }],
      apply: jest.fn(async () => ({
        amendmentsApplied: 1,
        commentsInserted: 0,
        errors: ['Chunk c1: original content no longer matches the staged range'],
      })),
      discard: jest.fn(async () => {}),
      failedCount: 0,
      cancelledCount: 0,
    };
    const actions = makeActions({ runDocumentSkill: jest.fn(async () => staged) });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));

    expect(cardEl.classList.contains('proposal-applied')).toBe(true);
    expect(view._msg.setStatus).toHaveBeenCalledWith(expect.stringContaining('1 section(s) skipped'));
  });

  test('format card warns when no formatting targets matched', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      applyFormatProposal: jest.fn(async () => ({ applied: false, appliedRanges: 0, insertedParagraphs: 0 })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'some selected text',
    });

    await conv.submit('把这段话加粗并标红');
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));

    expect(cardEl.classList.contains('proposal-applied')).toBe(false);
    expect(cardEl.classList.contains('proposal-warning')).toBe(true);
    expect(cardEl.querySelector('.proposal-card-status').textContent)
      .toContain('no formatting targets matched');
  });

  test('selection edit card shows an inline diff of the rewrite', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'some selected text',
    });

    await conv.submit('润色这段话');

    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    expect(cardEl.querySelectorAll('.proposal-card-change')).toHaveLength(1);
    expect(cardEl.querySelector('.diff-view del').textContent).toBe('before');
    expect(cardEl.querySelector('.diff-view ins').textContent).toBe('after');
  });

  test('onTurnCommitted fires after a turn completes with the live session', async () => {
    const appState = makeAppState();
    const view = makeView();
    const onTurnCommitted = jest.fn();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions: makeActions(), onTurnCommitted,
      getSelectionText: async () => '',
    });

    await conv.submit('what is the deadline?');

    expect(view._msg.finalizeForHistory).toHaveBeenCalledTimes(1);
    expect(onTurnCommitted).toHaveBeenCalledTimes(1);
    expect(onTurnCommitted).toHaveBeenCalledWith(
      expect.objectContaining({ id: 's-test' })
    );
  });

  test('attachProposal receives the full card plus a tracking meta object', async () => {
    const appState = makeAppState();
    const view = makeView();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions: makeActions(), getSelectionText: async () => 'some text',
    });

    await conv.submit('polish this');

    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    const [card, meta] = view._msg.attachProposal.mock.calls[0];
    expect(card.el).toBeDefined();
    expect(meta).toEqual(expect.objectContaining({
      title: 'Proposed edit',
      state: 'pending',
      items: expect.any(Array),
    }));
    // Note: the wrap-into-meta behavior lives in chat-view (covered there);
    // the mock here just records the call args.
  });
});

describe('createConversation.newChat', () => {
  test('persists the outgoing session before clearing', async () => {
    const appState = makeAppState();
    const view = makeView();
    const onTurnCommitted = jest.fn();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions: makeActions(), onTurnCommitted, getSelectionText: async () => '',
    });

    await conv.submit('what is the deadline?');
    expect(onTurnCommitted).toHaveBeenCalledTimes(1);

    conv.newChat();

    // Called once on submit, once on newChat for the outgoing session.
    expect(onTurnCommitted).toHaveBeenCalledTimes(2);
    expect(view.clearChat).toHaveBeenCalledTimes(1);
    expect(view.renderWelcome).toHaveBeenCalledTimes(1);
  });
});

describe('createConversation.cancel', () => {
  test('aborts the document and chat controllers and releases the UI immediately', () => {
    const docController = new AbortController();
    const chatController = new AbortController();
    const appState = makeAppState({
      processDocController: docController,
      chatController,
    });
    const input = makeInput();
    const conv = createConversation({
      appState, view: makeView(), input, log: jest.fn(), actions: makeActions(),
    });

    conv.cancel();

    // The captured controller reference retains its (now aborted) signal so
    // the assertion still proves cancellation propagated to the LLM layer.
    expect(docController.signal.aborted).toBe(true);
    expect(chatController.signal.aborted).toBe(true);
    // Cancel also frees the UI immediately so the user can interact again
    // without waiting for the in-flight fetch to settle. The orphan promise
    // guards its own finally on controller identity (see runDocumentTurn).
    expect(appState.processDocController).toBeNull();
    expect(appState.isProcessingDoc).toBe(false);
    expect(input.setProcessing).toHaveBeenCalledWith(false);
  });

  test('is a no-op when no controllers are active', () => {
    const appState = makeAppState();
    const input = makeInput();
    const conv = createConversation({
      appState, view: makeView(), input, log: jest.fn(), actions: makeActions(),
    });

    conv.cancel();

    expect(appState.processDocController).toBeNull();
    expect(appState.isProcessingDoc).toBe(false);
    expect(input.setProcessing).not.toHaveBeenCalled();
  });

  test('orphan settle after cancel does not clobber a follow-up turn', async () => {
    // A slow runDocumentSkill whose promise resolves AFTER cancel() — the
    // gating in runDocumentTurn's finally must keep that promise's finally
    // from releasing the new turn's state.
    let resolveSkill;
    const skillPromise = new Promise((resolve) => { resolveSkill = resolve; });
    const appState = makeAppState();
    const view = makeView();
    const input = makeInput();
    const actions = makeActions({
      runDocumentSkill: jest.fn(async () => skillPromise),
    });
    const conv = createConversation({
      appState, view, input, log: jest.fn(), actions,
    });

    // Kick off a document-scope turn (free-text DOC_EDIT intent → runDocumentTurn).
    const inFlight = conv.submit('please polish the document');

    // Wait a microtask so runDocumentTurn has set its controller, then cancel.
    await Promise.resolve();
    expect(appState.processDocController).not.toBeNull();
    const orphanedController = appState.processDocController;

    conv.cancel();
    expect(appState.processDocController).toBeNull();
    expect(appState.isProcessingDoc).toBe(false);
    expect(input.setProcessing).toHaveBeenLastCalledWith(false);

    // Simulate a follow-up turn that re-acquires the document flags.
    appState.isProcessingDoc = true;
    appState.processDocController = new AbortController();
    input.setProcessing(true);

    // Now let the orphan's runDocumentSkill resolve. Its finally must NOT
    // touch the new turn's flags.
    resolveSkill({ results: [], applicationResult: { amendmentsApplied: 0, commentsInserted: 0 }, chunks: [], cancelled: true });
    await inFlight;

    expect(appState.processDocController).not.toBe(orphanedController);
    expect(appState.processDocController).not.toBeNull();
    expect(appState.isProcessingDoc).toBe(true);
    // setProcessing was last called with true by the follow-up; the orphan
    // must not have flipped it back to false.
    expect(input.setProcessing).toHaveBeenLastCalledWith(true);
  });

  test('stop aborts an in-flight selection-edit LLM call and releases the UI', async () => {
    const appState = makeAppState();
    const view = makeView();
    const input = makeInput();
    const actions = makeActions({
      prepareSelectionAmendment: jest.fn((deps, args) => new Promise((_resolve, reject) => {
        args.signal.addEventListener('abort', () =>
          reject(new DOMException('The operation was aborted.', 'AbortError')));
      })),
    });
    const conv = createConversation({
      appState, view, input, log: jest.fn(), actions,
      getSelectionText: async () => 'some selected text',
    });

    const inFlight = conv.submit('make it formal');
    await Promise.resolve();
    await Promise.resolve();
    const controller = appState.chatController;
    expect(controller).not.toBeNull();

    conv.cancel();
    expect(controller.signal.aborted).toBe(true);

    await inFlight;
    expect(view._msg.setStatus).toHaveBeenCalledWith('Cancelled.');
    expect(view._msg.markError).not.toHaveBeenCalled();
    // No proposal card from a cancelled draft.
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(appState.isProcessing).toBe(false);
    expect(appState.chatController).toBeNull();
    expect(input.setProcessing).toHaveBeenLastCalledWith(false);
  });

  test('stop during a compound turn aborts the current task and skips the rest', async () => {
    const appState = makeAppState();
    const view = makeView();
    const input = makeInput();
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({
        tasks: [
          { type: 'format', instruction: '增加标题' },
          { type: 'format', instruction: '居中全文' },
        ],
        model: 'm',
      })),
      // Task one hangs until its signal aborts — as a real stream would.
      prepareFormatProposal: jest.fn((deps, args) => new Promise((_resolve, reject) => {
        args.signal.addEventListener('abort', () =>
          reject(new DOMException('The operation was aborted.', 'AbortError')));
      })),
    });
    const conv = createConversation({
      appState, view, input, log: jest.fn(), actions,
      getSelectionText: async () => '',
    });

    const inFlight = conv.submit('增加标题，并深度润色修改');
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(appState.chatController).not.toBeNull();

    conv.cancel();
    await inFlight;

    // Task one started (and was aborted); task two never started.
    expect(actions.prepareFormatProposal).toHaveBeenCalledTimes(1);
    expect(view._msg.setStatus).toHaveBeenCalledWith('Cancelled.');
    expect(appState.isProcessing).toBe(false);
    expect(appState.chatController).toBeNull();
    expect(input.setProcessing).toHaveBeenLastCalledWith(false);
  });
});

describe('chunkCitation', () => {
  test('label is the first ~6 words of the first non-empty paragraph', () => {
    const c = chunkCitation({ id: 'c1', paragraphs: [{ text: '' }, { text: 'one two three four five six seven eight' }] });
    expect(c.label).toBe('one two three four five six');
    expect(c.searchText).toBe('one two three four five six seven eight');
  });

  test('falls back to the chunk id when no paragraph text exists', () => {
    const c = chunkCitation({ id: 'c2', paragraphs: [] });
    expect(c.label).toBe('c2');
  });
});
