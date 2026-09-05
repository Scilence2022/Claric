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
const { BUILTIN_SKILLS, listSkills } = require('../src/taskpane/skills.js');

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
    readDocumentTableRegions: jest.fn(async () => [{
      tableIndex: 1,
      rowCount: 2,
      colCount: 1,
      values: [['a'], ['b']],
      bounds: { startRow: 1, endRow: 2, startCol: 1, endCol: 1 },
      cells: [{ row: 1, col: 1, text: 'a' }, { row: 2, col: 1, text: 'b' }],
      merged: false,
      style: null,
    }]),
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

  test('figure legend review-edit requests use the selected image tool session', () => {
    const facts = {
      hasSelection: true, hasImageSelection: true, hasTextSelection: false,
      skills: BUILTIN_SKILLS,
    };
    expect(routeTurn('选择图像的 Figure legends 是否有改进的空间？如果有请改进', facts).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('Review this Figure caption and improve it if needed', facts).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
  });

  test('mixed image+text selection keeps text routing except for figure caption work', () => {
    const mixed = {
      hasSelection: true, hasImageSelection: true, hasTextSelection: true,
      skills: BUILTIN_SKILLS,
    };
    expect(routeTurn('make it formal', mixed).type).toBe(TURN_TYPE.SELECTION_EDIT);
    expect(routeTurn('what does this say?', mixed).type).toBe(TURN_TYPE.DOC_QA);
    expect(routeTurn('评估图注，如有需要请改进', mixed).type).toBe(TURN_TYPE.IMAGE_TOOL);

    const textOnly = {
      hasSelection: true, hasImageSelection: false, hasTextSelection: true,
      skills: BUILTIN_SKILLS,
    };
    expect(routeTurn('改进这段 Figure caption', textOnly).type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('image-only selection takes document scope for format intent', () => {
    const turn = routeTurn('把标题居中', { hasSelection: true, hasImageSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.FORMAT);
    expect(turn.scope).toBe('document');
    // Text selection keeps selection scope.
    expect(routeTurn('把标题居中', { hasSelection: true, skills: BUILTIN_SKILLS }).scope).toBe('selection');
  });

  test('multi-cell table selection routes to the table tool session', () => {
    // Edit-ish instruction: text pipelines can't represent cell boundaries.
    expect(routeTurn('润色表格', {
      hasSelection: true, hasMultiCellTableRegion: true, skills: BUILTIN_SKILLS,
    }).type).toBe(TURN_TYPE.TABLE_TOOL);
    // Questions too: visual reading is get_state, not injected markdown.
    expect(routeTurn('这个表格有什么问题?', {
      hasSelection: true, hasMultiCellTableRegion: true, skills: BUILTIN_SKILLS,
    }).type).toBe(TURN_TYPE.TABLE_TOOL);
    expect(routeTurn('summarize this table', {
      hasSelection: true, hasMultiCellTableRegion: true, skills: BUILTIN_SKILLS,
    }).type).toBe(TURN_TYPE.TABLE_TOOL);
  });

  test('intra-cell text selection stays on the flat-text SELECTION_EDIT path', () => {
    // parentTableCell non-null → single cell — text pipeline.
    const turn = routeTurn('改一下这段', { hasSelection: true, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.SELECTION_EDIT);
  });

  test('table-look format intent with a table selection enters the table tool session', () => {
    // The instruction names the table's look (居中 the 表格) — the table tool
    // session's style tools own that, not the paragraph format pipeline.
    const turn = routeTurn('把表格居中', {
      hasSelection: true, hasMultiCellTableRegion: true, skills: BUILTIN_SKILLS,
    });
    expect(turn.type).toBe(TURN_TYPE.TABLE_TOOL);

    // Non-table format targets with a table selection stay FORMAT at
    // document scope (format ops target paragraphs, not table cells).
    const heading = routeTurn('把标题居中', {
      hasSelection: true, hasMultiCellTableRegion: true, skills: BUILTIN_SKILLS,
    });
    expect(heading.type).toBe(TURN_TYPE.FORMAT);
    expect(heading.scope).toBe('document');
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

  test('a Chinese yes/no question with 吗 routes to Q&A, not FORMAT', () => {
    // "Can you change the style of the selected table?" — contains 样式
    // (format intent) but is a question, so it must be ANSWERED, not parsed
    // as a format-op request that yields "no changes".
    const turn = routeTurn('你能改变选择的表格的样式吗？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
    expect(turn.question).toBe('你能改变选择的表格的样式吗？');
    expect(turn.type).not.toBe(TURN_TYPE.FORMAT);
  });

  test('a trailing question mark routes to Q&A even without a question lead', () => {
    const turn = routeTurn('样式能改吗？', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.DOC_QA);
    expect(turn.type).not.toBe(TURN_TYPE.FORMAT);
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

  test('image styling intents (align/link/alt-title/scale) enter the image tool loop', () => {
    expect(routeTurn('把图片居中', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('给第一张图片加超链接 https://example.com', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('改一下图片的标题', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('把图片缩放到一半', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('center the image', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
  });

  test('image questions stay Q&A even with management verbs', () => {
    expect(routeTurn('如何删除文档里的图片？', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOC_QA);
  });

  test('document-scope image/table intents route to their tool sessions', () => {
    // Plural-marked document intents, unambiguous across intent families:
    // every picture / every table.
    expect(routeTurn('给所有图片都加上 alt 文字', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOCUMENT_IMAGE_TOOL);
    expect(routeTurn('把图片都加上 alt 文字', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOCUMENT_IMAGE_TOOL);
    expect(routeTurn('给所有表格加边框', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOCUMENT_TABLE_TOOL);
    expect(routeTurn('every table to three-line', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.DOCUMENT_TABLE_TOOL);

    // Ambiguous phrasings that ALSO hit the format family (标题/居中/样式)
    // compound — the planner decides; the doc turn never fires alone.
    expect(routeTurn('把图片都加上标题', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.COMPOUND);
    expect(routeTurn('all images centered', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.COMPOUND);

    // Negative cases: no plural marker → single-object routing stays put.
    expect(routeTurn('把图片居中', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .toBe(TURN_TYPE.IMAGE_TOOL);
    expect(routeTurn('把表格改成三线表', { hasSelection: false, skills: BUILTIN_SKILLS }).type)
      .not.toBe(TURN_TYPE.DOCUMENT_TABLE_TOOL);
  });

  test('document image/table intents compound with text-edit families', () => {
    // "润色全文 + 给所有图片加 alt 文字" → COMPOUND (edit + image_management).
    const turn = routeTurn('润色全文，给所有图片加 alt 文字', { hasSelection: false, skills: BUILTIN_SKILLS });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);

    // A multi-cell table selection already routes to TABLE_TOOL — no
    // planner trip for a single instruction on a selected table.
    const selected = routeTurn('把表格都改成三线表', {
      hasSelection: true, hasMultiCellTableRegion: true, skills: BUILTIN_SKILLS,
    });
    expect(selected.type).toBe(TURN_TYPE.TABLE_TOOL);
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

  test('image replace op stages a before/after visual diff (no top preview)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareImageToolEdit: jest.fn(async () => ({
        instruction: 'x',
        ops: [{
          type: 'replace', index: 1, instruction: 'improve legends',
          svg: '<svg width="10" height="10"><rect width="10" height="10"/></svg>',
          beforeSrc: 'data:image/png;base64,iVBORw0KGgo=',
        }],
        items: [{
          id: 1, label: 'Replace image 1',
          before: 'existing picture', after: 'improve legends (0.1 KB SVG → PNG)',
          svg: '<svg width="10" height="10"><rect width="10" height="10"/></svg>',
          beforeSrc: 'data:image/png;base64,iVBORw0KGgo=',
        }],
        snapshotCount: 1,
        model: 'm',
        toolLoop: { steps: 2, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('改进文档里第二张图片的图例');

    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    // The replace item renders its own two-up diff; the card-level preview
    // (single-svg insert convention) stays empty for replace ops.
    expect(cardEl.querySelector(':scope > .proposal-card-preview-svg')).toBeNull();
    const diff = cardEl.querySelector('.proposal-card-image-diff');
    expect(diff).not.toBeNull();
    expect(diff.querySelector('img').getAttribute('src')).toBe('data:image/png;base64,iVBORw0KGgo=');
    expect(diff.querySelector('svg')).not.toBeNull();
  });

  test('a single insert op keeps the top-of-card preview (no per-item duplicate)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const svg = '<svg width="10" height="10"><rect width="10" height="10"/></svg>';
    const actions = makeActions({
      prepareImageToolEdit: jest.fn(async () => ({
        instruction: 'x',
        ops: [{ type: 'insert', position: 'end', instruction: 'a sun', svg }],
        items: [{ id: 1, label: 'Insert illustration at end', before: '', after: 'a sun (0.1 KB SVG → PNG)', svg }],
        snapshotCount: 0,
        model: 'm',
        toolLoop: { steps: 1, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('加一张太阳的插图');

    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    expect(cardEl.querySelector(':scope > .proposal-card-preview-svg svg')).not.toBeNull();
    // The svg rides the top preview only — not duplicated inside the row.
    expect(cardEl.querySelector('.proposal-card-image-diff')).toBeNull();
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

  test('mixed image+caption selection passes both contexts to the image tool', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions,
      getSelectionContent: async () => ({
        text: 'Figure 2. Existing caption',
        totalImages: 1,
        images: [{
          dataUrl: 'data:image/png;base64,X', width: 300, height: 200,
          altText: 'architecture figure', identityKey: 'image-key-2',
        }],
      }),
    });

    await conv.submit('选择图像的 Figure legends 是否有改进的空间？如果有请改进');

    expect(actions.prepareImageToolEdit).toHaveBeenCalledTimes(1);
    const args = actions.prepareImageToolEdit.mock.calls[0][1];
    expect(args.selectionText).toBe('Figure 2. Existing caption');
    expect(args.selectionImages).toEqual([{
      width: 300, height: 200, altText: 'architecture figure', identityKey: 'image-key-2',
    }]);
    expect(actions.prepareSelectionAmendment).not.toHaveBeenCalled();
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

  test('multi-cell table selection enters the table tool session with ops', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareTableToolEdit: jest.fn(async () => ({
        instruction: '给第一行加一行合计',
        selectionText: 'R1C1 R1C2',
        amendedText: null,
        commentText: null,
        model: 'm',
        tablePatch: {
          rowCount: 3, colCount: 2,
          cells: [{ row: 3, col: 1, text: '合计' }],
          rowOps: [],
          bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
          originals: [['A', 'B'], ['a', 'b'], ['合计', '']],
        },
        tableItems: [{ id: 0, label: 'Cell R3C1', before: '', after: '合计', searchText: undefined }],
        toolLoop: { steps: 3, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions,
      getSelectionContent: async () => ({
        text: 'R1C1 R1C2',
        hasMultiCellTableRegion: true,
        tableRegion: { startRow: 1, endRow: 1, startCol: 1, endCol: 2 },
      }),
    });

    await conv.submit('给第一行加一行合计');

    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(actions.prepareTableToolEdit.mock.calls[0][1].instruction).toBe('给第一行加一行合计');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
  });

  test('multi-cell table selection + analysis question routes to the loop (noOps answer)', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareTableToolEdit: jest.fn(async () => ({
        noOps: true,
        answer: 'two columns, totals match header sum',
        selectionText: 'header\nrow a\nrow b',
        model: 'm',
        toolLoop: { steps: 1, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions,
      getSelectionContent: async () => ({
        text: 'header\nrow a\nrow b',
        hasMultiCellTableRegion: true,
        tableRegion: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
      }),
    });

    await conv.submit('总结一下这个表格');

    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(view._msg.setText).toHaveBeenCalledWith('two columns, totals match header sum');
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
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

  test('all-chunks-failed staged run reports the failure with a retry link, not "no changes"', async () => {
    const appState = makeAppState();
    const view = makeView();
    const retryFailed = jest.fn(async () => {});
    const staged = {
      staged: true,
      results: [
        { status: 'rejected', amendment: null, error: 'HTTP 500: backend down', chunk: { id: 'c0' } },
        { status: 'rejected', amendment: null, error: 'HTTP 500: backend down', chunk: { id: 'c1' } },
      ],
      chunks: [{ id: 'c0' }, { id: 'c1' }],
      apply: jest.fn(),
      discard: jest.fn(async () => {}),
      retryFailed,
      failedCount: 2,
      cancelledCount: 0,
    };
    const actions = makeActions({
      runDocumentSkill: jest.fn(async () => staged),
    });
    const logWithRetry = jest.fn();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(), logWithRetry,
      actions, getSelectionText: async () => '',
    });

    await conv.submit('please polish the whole document');

    // The systemic failure must not be masqueraded as an empty proposal.
    expect(staged.discard).not.toHaveBeenCalled();
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(view._msg.setStatus).toHaveBeenCalledWith(expect.stringContaining('failed on 2 of 2 section(s)'));
    expect(view._msg.setStatus).toHaveBeenCalledWith(expect.stringContaining('backend down'));
    // The retry link drives the outcome's retryFailed handle.
    expect(logWithRetry).toHaveBeenCalledWith(
      expect.stringContaining('retry'),
      'warning',
      expect.any(Function)
    );
    logWithRetry.mock.calls[0][2]();
    expect(retryFailed).toHaveBeenCalledTimes(1);
    expect(staged.apply).not.toHaveBeenCalled();
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

  test('parameterized new summary skill passes its focus to the summary pipeline', async () => {
    const appState = makeAppState();
    const actions = makeActions();
    const conv = createConversation({
      appState, view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/executive-summary focus on financial risks');

    expect(actions.runSummarySkill).toHaveBeenCalledTimes(1);
    const { promptTemplate } = actions.runSummarySkill.mock.calls[0][1];
    expect(promptTemplate).toContain('Write an executive summary for a busy decision-maker');
    expect(promptTemplate).toContain('Additional instructions from the user: focus on financial risks');
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

  test('parameterized new amendment skill stages a selection proposal with its instructions', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'This selected passage needs editing.',
    });

    await conv.submit('/polish preserve field-specific terminology');

    expect(actions.prepareSelectionAmendment).toHaveBeenCalledTimes(1);
    const { promptTemplate } = actions.prepareSelectionAmendment.mock.calls[0][1];
    expect(promptTemplate).toContain('Revise {selection} into polished academic English');
    expect(promptTemplate).toContain('Additional instructions from the user: preserve field-specific terminology');
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(1);
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
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

  test('parameterized new chat skill returns key points through document Q&A', async () => {
    const appState = makeAppState();
    const actions = makeActions();
    const conv = createConversation({
      appState, view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('/key-points emphasize deadlines and owners');

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    const qaArgs = actions.answerQuestion.mock.calls[0][1];
    expect(qaArgs.question).toBe('emphasize deadlines and owners');
    expect(qaArgs.skillTemplate).toContain('most important points in the current document');
    expect(actions.runSummarySkill).not.toHaveBeenCalled();
  });

  test('parameterized comment skill routes and executes through the comment pipeline', async () => {
    const appState = makeAppState();
    const actions = makeActions();
    const conv = createConversation({
      appState, view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'The results are clear.',
    });

    await conv.submit('/check-clarity focus on ambiguous references');

    expect(actions.fireSelectionComment).toHaveBeenCalledTimes(1);
    expect(actions.fireSelectionComment.mock.calls[0][1].promptTemplate).toContain('Review {selection} for clarity');
    expect(actions.fireSelectionComment.mock.calls[0][1].promptTemplate).toContain('Additional instructions from the user: focus on ambiguous references');
    expect(actions.runDocumentSkill).not.toHaveBeenCalled();
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

    // The staged card carries an inline SVG preview of the proposed artwork
    // (inline render — SVG data URLs fail to decode on some hosts).
    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    const preview = cardEl.querySelector('.proposal-card-preview-svg');
    expect(preview).not.toBeNull();
    expect(preview.querySelector('svg')).not.toBeNull();
    expect(preview.querySelector('script')).toBeNull();

    cardEl.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(actions.applyIllustrationProposal).toHaveBeenCalledTimes(1);
  });

  test('illustration turn stages a raster preview with its supplied MIME type', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareIllustrationProposal: jest.fn(async () => ({
        instruction: 'x', svg: null, imageBase64: '/9j/4AAQ',
        previewSrc: 'data:image/jpeg;base64,/9j/4AAQ',
        position: 'end', model: 'image-model', sizeLabel: '1 KB image',
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('设计示意图并插入');

    const cardEl = view._msg.attachProposal.mock.calls[0][0].el;
    const preview = cardEl.querySelector('img.proposal-card-preview');
    expect(preview).not.toBeNull();
    expect(preview.getAttribute('src')).toBe('data:image/jpeg;base64,/9j/4AAQ');
    expect(actions.applyIllustrationProposal).not.toHaveBeenCalled();
  });

  test('illustration turn with no usable SVG shows a status instead of a card', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareIllustrationProposal: jest.fn(async () => ({ instruction: 'x', svg: null, imageBase64: null, position: 'end', model: 'm' })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('设计并增加SVG插图');

    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    // Message is engine-agnostic now that an illustration may come from the
    // image model (imageBase64) or the SVG route (svg).
    expect(view._msg.setStatus).toHaveBeenCalledWith('The model produced no usable illustration.');
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

  test('compound turn dispatches image_management/table_management to document-scope tool sessions', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({
        tasks: [
          { type: 'image_management', instruction: '给所有图片加上标题' },
          { type: 'table_management', instruction: '把表格改成三线表样式' },
        ],
        model: 'm',
      })),
      prepareImageToolEdit: jest.fn(async () => ({
        instruction: 'x',
        ops: [{ type: 'altText', index: 1, text: '标题' }],
        items: [{ id: 1, label: 'Alt text for image 1', before: '', after: '标题' }],
        snapshotCount: 3,
        model: 'm',
        toolLoop: { steps: 2, finished: true },
      })),
      prepareTableToolEdit: jest.fn(async () => ({
        selectionText: 'a\nb',
        amendedText: null,
        commentText: null,
        model: 'm',
        tablePatch: {
          rowCount: 2, colCount: 1, cells: [], rowOps: [], merges: [],
          styleOps: [{ type: 'borders', borders: { all: { type: 'none' } } }],
          bounds: { startRow: 1, endRow: 2, startCol: 1, endCol: 1 },
          originals: [['a'], ['b']],
        },
        tableItems: [{ label: 'Borders: all none' }],
        toolLoop: { steps: 3, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(),
      actions,
      getSelectionText: async () => '',
      getSelectionImages: async () => [],
    });

    await conv.submit('润色全文，给所有图片加标题，把表格改成三线表');

    // image_management runs the whole-document image tool session (no
    // selectionImages — snapshot covers every picture).
    expect(actions.prepareImageToolEdit).toHaveBeenCalledTimes(1);
    expect(actions.prepareImageToolEdit.mock.calls[0][1].instruction).toBe('给所有图片加上标题');
    expect(actions.prepareImageToolEdit.mock.calls[0][1].selectionImages).toBeUndefined();
    // table_management runs the table tool session against the document's
    // table REGIONS (all tables by tableIndex), producing its own card.
    expect(actions.readDocumentTableRegions).toHaveBeenCalledTimes(1);
    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(actions.prepareTableToolEdit.mock.calls[0][1].regions).toBeInstanceOf(Array);
    expect(actions.prepareTableToolEdit.mock.calls[0][1].regions[0].tableIndex).toBe(1);
    // One card per task; applies are gated behind the cards.
    expect(view._msg.attachProposal).toHaveBeenCalledTimes(2);
    expect(actions.applyImageOps).not.toHaveBeenCalled();
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
    expect(staged.apply).toHaveBeenCalledWith(['c0'], expect.anything());
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

describe('createConversation.submit with file attachments', () => {
  test('text attachments append labeled sections to the QA question', async () => {
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState: makeAppState(), view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('what do these files say?', [
      { name: 'a.txt', kind: 'text', size: 9, text: 'alpha body' },
      { name: 'b.pdf', kind: 'pdf', size: 8, text: 'beta body' },
    ]);

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    const args = actions.answerQuestion.mock.calls[0][1];
    expect(args.question).toContain('what do these files say?');
    expect(args.question).toContain('--- ATTACHED FILE: a.txt ---\nalpha body');
    expect(args.question).toContain('--- ATTACHED FILE: b.pdf ---\nbeta body');
    // The user bubble shows the typed text plus display metadata only.
    expect(view.addUserMessage).toHaveBeenCalledWith('what do these files say?', [
      { name: 'a.txt', kind: 'text', size: 9 },
      { name: 'b.pdf', kind: 'pdf', size: 8 },
    ]);
  });

  test('image attachments ride the QA turn as image parts', async () => {
    const actions = makeActions();
    const conv = createConversation({
      appState: makeAppState(), view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('what is in this picture?', [
      { name: 'p.png', kind: 'image', size: 3, dataUrl: 'data:image/png;base64,AQID' },
    ]);

    const args = actions.answerQuestion.mock.calls[0][1];
    expect(args.questionImages).toEqual([{ name: 'p.png', dataUrl: 'data:image/png;base64,AQID' }]);
    expect(args.question).toContain('ATTACHED IMAGE: p.png');
    expect(args.question).not.toContain('data:image');
  });

  test('attachments alone (no text) submit with a default instruction', async () => {
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState: makeAppState(), view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('   ', [{ name: 'a.txt', kind: 'text', size: 5, text: 'body' }]);

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    expect(actions.answerQuestion.mock.calls[0][1].question).toContain('What do the attached file(s) say?');
    expect(view.addUserMessage.mock.calls[0][0]).toBe('What do the attached file(s) say?');
  });

  test('empty text and no attachments is still a no-op', async () => {
    const view = makeView();
    const actions = makeActions();
    const conv = createConversation({
      appState: makeAppState(), view, input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => '',
    });

    await conv.submit('   ', []);
    await conv.submit('', undefined);

    expect(actions.answerQuestion).not.toHaveBeenCalled();
    expect(view.addUserMessage).not.toHaveBeenCalled();
  });

  test('attachment text appends to edit instructions too', async () => {
    const actions = makeActions();
    const conv = createConversation({
      appState: makeAppState(), view: makeView(), input: makeInput(), log: jest.fn(),
      actions, getSelectionText: async () => 'clause text',
    });

    await conv.submit('make it formal', [{ name: 'style.md', kind: 'text', size: 4, text: 'rules' }]);

    expect(actions.prepareSelectionAmendment).toHaveBeenCalledTimes(1);
    expect(actions.prepareSelectionAmendment.mock.calls[0][1].promptTemplate)
      .toContain('--- ATTACHED FILE: style.md ---\nrules');
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
  test('aborts controllers without releasing a still-settling write lock', () => {
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
    expect(appState.processDocController).toBe(docController);
    expect(input.setProcessing).not.toHaveBeenCalledWith(false);
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
    expect(appState.processDocController).toBe(orphanedController);
    expect(appState.isProcessingDoc).toBe(true);
    expect(input.setProcessing).toHaveBeenLastCalledWith(true);
    await conv.submit('what is this?');
    expect(actions.answerQuestion).not.toHaveBeenCalled();

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

describe('/mcp reserved skill routing', () => {
    test('routeTurn sends /mcp to the reserved tools skill', () => {
        const turn = routeTurn('/mcp search the web for X', {
            hasSelection: false,
            skills: listSkills(makeAppState().promptManager),
        });
        expect(turn.type).toBe(TURN_TYPE.SKILL);
        expect(turn.skill).toEqual(expect.objectContaining({ name: 'mcp', category: 'tools', reserved: true }));
        expect(turn.args).toBe('search the web for X');
    });

    test('with no MCP servers configured the turn explains itself instead of failing', async () => {
        const appState = makeAppState({ config: {} });
        const view = makeView();
        const conv = createConversation({
            appState, view, input: makeInput(), log: jest.fn(),
            actions: makeActions(), getSelectionText: async () => '',
        });

        await conv.submit('/mcp hello');
        expect(view._msg.appendText).toHaveBeenCalledWith(expect.stringContaining('No MCP servers are configured'));
    });
});

describe('turn lifecycle guards', () => {
  test('a chained-instruction tool loop with a read-only outcome answers in chat (no card)', async () => {
    // prepareTableToolEdit can resolve to {noOps, answer} — no tablePatch and
    // no amendedText. Staging a card for it produced an apply that wrote
    // nothing yet reported "Applied as tracked changes".
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      prepareTableToolEdit: jest.fn(async () => ({
        noOps: true,
        answer: 'the table already totals correctly',
        selectionText: 'header\nrow a',
        model: 'm',
        toolLoop: { steps: 2, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(), actions,
      getSelectionText: async () => 'header\nrow a',
    });

    // "然后" makes this a chained instruction, which takes the tool loop first.
    await conv.submit('检查合计，然后修正表头');

    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(view._msg.setText).toHaveBeenCalledWith('the table already totals correctly');
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
    expect(actions.applySelectionAmendment).not.toHaveBeenCalled();
  });

  test('a cleanup turn registers its controller so cancel() reaches it', async () => {
    const appState = makeAppState();
    const view = makeView();
    const input = makeInput();
    const actions = makeActions({
      prepareEmptyParagraphCleanup: jest.fn(() => new Promise(() => {})),
    });
    const conv = createConversation({
      appState, view, input, log: jest.fn(), actions,
      getSelectionText: async () => '',
    });

    conv.submit('删除多余的空行');
    await Promise.resolve();
    await Promise.resolve();

    expect(appState.chatController).not.toBeNull();
    conv.cancel();
    expect(appState.chatController.signal.aborted).toBe(true);
  });

  test('a compound sub-task keeps the parent controller and busy flag', async () => {
    // Sub-runners share the compound turn's controller: one must not null it
    // (nor drop the busy flag) when it finishes, or cancel() loses its grip on
    // the rest of the chain and a later turn can start mid-chain.
    const appState = makeAppState();
    const view = makeView();
    const input = makeInput();
    const seen = [];
    const actions = makeActions({
      planDocumentTasks: jest.fn(async () => ({
        tasks: [
          { type: 'edit', instruction: '删除多余的空行' },   // → CLEANUP
          { type: 'format', instruction: '居中全文' },
        ],
        model: 'm',
      })),
      prepareEmptyParagraphCleanup: jest.fn(async () => ({ emptyCount: 0 })),
      prepareFormatProposal: jest.fn(async () => {
        seen.push({
          controller: appState.chatController,
          isProcessing: appState.isProcessing,
        });
        return { instruction: 'x', scope: 'document', ops: [{ font: { bold: true } }], model: 'm' };
      }),
    });
    const conv = createConversation({
      appState, view, input, log: jest.fn(), actions,
      getSelectionText: async () => '',
    });

    const inFlight = conv.submit('删除多余的空行，并居中全文');
    await Promise.resolve();
    await Promise.resolve();
    const compoundController = appState.chatController;
    await inFlight;

    // The cleanup sub-task ran first; when task two started, the compound
    // turn's controller was still registered and the pane still busy.
    expect(actions.prepareEmptyParagraphCleanup).toHaveBeenCalledTimes(1);
    expect(seen).toHaveLength(1);
    expect(seen[0].controller).toBe(compoundController);
    expect(seen[0].isProcessing).toBe(true);
    // The whole chain settled: flag down, controller released.
    expect(appState.isProcessing).toBe(false);
    expect(appState.chatController).toBeNull();
    expect(input.setProcessing).toHaveBeenLastCalledWith(false);
  });

  test('planning failure re-routes with the table-region flag intact', async () => {
    const appState = makeAppState();
    const view = makeView();
    const actions = makeActions({
      // Planner returns nothing → the fallback re-routes the instruction.
      planDocumentTasks: jest.fn(async () => ({ tasks: [], model: 'm' })),
      prepareTableToolEdit: jest.fn(async () => ({
        noOps: true, answer: 'reviewed', selectionText: 'x', model: 'm',
        toolLoop: { steps: 1, finished: true },
      })),
    });
    const conv = createConversation({
      appState, view, input: makeInput(), log: jest.fn(), actions,
      getSelectionContent: async () => ({
        text: 'header\nrow a',
        hasMultiCellTableRegion: true,
        images: [],
      }),
    });

    // Format + edit intents → COMPOUND; planning fails → fallback routing.
    // With the flag carried, the fallback reaches TABLE_TOOL; with it dropped
    // the same instruction re-routed to the FORMAT pipeline instead.
    await conv.submit('表头加粗，并润色');

    expect(actions.planDocumentTasks).toHaveBeenCalledTimes(1);
    expect(actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(actions.prepareFormatProposal).not.toHaveBeenCalled();
    expect(view._msg.setText).toHaveBeenCalledWith('reviewed');
  });

  test('routeTurn carries the table-region flag on compound turns', () => {
    // Multi-intent instruction with a multi-cell table selection: the planner
    // decides, and the flag must ride along so a planning failure can
    // re-route with the same selection facts runCompoundTurn was given.
    const turn = routeTurn('增加标题，并深度润色修改', {
      hasSelection: true,
      hasMultiCellTableRegion: true,
      skills: BUILTIN_SKILLS,
    });
    expect(turn.type).toBe(TURN_TYPE.COMPOUND);
    expect(turn.hasMultiCellTableRegion).toBe(true);

    // Without a table selection the flag is false, not absent.
    expect(routeTurn('增加标题，并深度润色修改', { hasSelection: false, skills: BUILTIN_SKILLS })
      .hasMultiCellTableRegion).toBe(false);
  });

  test('takes ownership synchronously before selection read and drops a new-chat late read', async () => {
    let resolveSelection;
    const getSelectionText = jest.fn(() => new Promise((resolve) => { resolveSelection = resolve; }));
    const view = makeView(); const input = makeInput(); const appState = makeAppState();
    const actions = makeActions();
    const conv = createConversation({ appState, view, input, actions, getSelectionText, log: jest.fn() });
    const first = conv.submit('what is this?');
    expect(appState.isProcessing).toBe(true);
    await conv.submit('what is that?');
    expect(getSelectionText).toHaveBeenCalledTimes(1);
    conv.newChat();
    resolveSelection('old selection');
    await first;
    expect(view.addUserMessage).not.toHaveBeenCalled();
    expect(actions.answerQuestion).not.toHaveBeenCalled();
    expect(appState.isProcessing).toBe(false);
  });

  test('late model output after newChat cannot finalize into the new session', async () => {
    let resolveAnswer;
    const actions = makeActions({ answerQuestion: jest.fn(() => new Promise((resolve) => { resolveAnswer = resolve; })) });
    const view = makeView(); const onTurnCommitted = jest.fn();
    const conv = createConversation({ appState: makeAppState(), view, input: makeInput(), actions,
      getSelectionText: async () => '', onTurnCommitted, log: jest.fn() });
    const first = conv.submit('what is this?');
    await Promise.resolve();
    conv.newChat();
    resolveAnswer('late answer');
    await first;
    expect(view._msg.setText).not.toHaveBeenCalledWith('late answer');
    expect(view._msg.finalizeForHistory).not.toHaveBeenCalled();
    expect(onTurnCommitted).toHaveBeenCalledTimes(1);
  });

  test.each(['format', 'image'])('%s card forwards its signal and holds the write lock until settlement', async (kind) => {
    let settle;
    let signal;
    const actionName = kind === 'format' ? 'applyFormatProposal' : 'applyImageOps';
    const actions = makeActions({ [actionName]: jest.fn((_deps, _proposal, ctx) => {
      signal = ctx.signal;
      return new Promise((resolve) => { settle = resolve; });
    }) });
    const appState = makeAppState(); const view = makeView(); const input = makeInput();
    const conv = createConversation({ appState, view, input, actions, log: jest.fn(), getSelectionText: async () => '' });
    await conv.submit(kind === 'format' ? 'bold the document' : 'delete images');
    const card = view._msg.attachProposal.mock.calls[0][0];
    card.el.querySelector('.btn-primary').click();
    await Promise.resolve();
    expect(signal).toBeInstanceOf(AbortSignal);
    conv.cancel();
    expect(signal.aborted).toBe(true);
    expect(appState.isProcessingDoc).toBe(true);
    await conv.submit('what is this?');
    expect(actions.answerQuestion).not.toHaveBeenCalled();
    settle({ interrupted: true, partial: true, appliedRanges: 1, insertedParagraphs: 0, applied: 1 });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(appState.isProcessingDoc).toBe(false);
    expect(card.el.textContent).toContain('Review the document');
  });

  test.each(['cancel', 'newChat'])('cleans an unattached late format proposal after %s', async (action) => {
    let resolve;
    const proposal = { anchor: { bookmark: '_late' }, ops: [{ font: { bold: true } }] };
    const actions = makeActions({ prepareFormatProposal: jest.fn(() => new Promise((done) => { resolve = done; })), discardFormatProposal: jest.fn() });
    const view = makeView();
    const conv = createConversation({ appState: makeAppState(), view, input: makeInput(), actions, log: jest.fn(), getSelectionText: async () => '' });
    const running = conv.submit('bold the document');
    await Promise.resolve();
    conv[action]();
    resolve(proposal);
    await running;
    expect(actions.discardFormatProposal).toHaveBeenCalledWith(expect.anything(), proposal);
    expect(view._msg.attachProposal).not.toHaveBeenCalled();
  });

  test('newChat defers format anchor disposal until the Word write settles', async () => {
    let settle;
    const actions = makeActions({ applyFormatProposal: jest.fn(() => new Promise((done) => { settle = done; })), discardFormatProposal: jest.fn() });
    const view = makeView();
    const conv = createConversation({ appState: makeAppState(), view, input: makeInput(), actions, log: jest.fn(), getSelectionText: async () => '' });
    await conv.submit('bold the document');
    const card = view._msg.attachProposal.mock.calls[0][0];
    card.el.querySelector('.btn-primary').click();
    conv.newChat();
    await Promise.resolve();
    expect(actions.discardFormatProposal).not.toHaveBeenCalled();
    settle({ appliedRanges: 1, insertedParagraphs: 0 });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(actions.discardFormatProposal).toHaveBeenCalledTimes(1);
  });

  test('newChat disposes pending illustration and blocks its old card', async () => {
    const actions = makeActions({ discardIllustrationProposal: jest.fn() });
    const view = makeView();
    const conv = createConversation({ appState: makeAppState(), view, input: makeInput(), actions, log: jest.fn(), getSelectionText: async () => '' });
    await conv.submit('draw an illustration');
    const card = view._msg.attachProposal.mock.calls[0][0];
    conv.newChat();
    await Promise.resolve();
    card.el.querySelector('.btn-primary').click();
    expect(actions.discardIllustrationProposal).toHaveBeenCalledTimes(1);
    expect(actions.applyIllustrationProposal).not.toHaveBeenCalled();
  });

  test('MCP cancellation reaches connect before listTools and skips subsequent servers', async () => {
    const mcp = require('../src/lib/mcp-client.js');
    let signal;
    const connect = jest.spyOn(mcp, 'connectMcpServer').mockImplementation((opts) => {
      signal = opts.signal;
      return new Promise((_resolve, reject) => signal.addEventListener('abort', () => reject(new DOMException('cancelled', 'AbortError'))));
    });
    try {
      const appState = makeAppState({ config: { mcpServers: [{ url: 'https://one.test' }, { url: 'https://two.test' }] } });
      const conv = createConversation({ appState, view: makeView(), input: makeInput(), actions: makeActions(), log: jest.fn(), getSelectionText: async () => '' });
      const running = conv.submit('/mcp list tools');
      await Promise.resolve();
      expect(signal).toBe(appState.chatController.signal);
      conv.cancel();
      await running;
      expect(signal.aborted).toBe(true);
      expect(connect).toHaveBeenCalledTimes(1);
      expect(appState.isProcessing).toBe(false);
    } finally { connect.mockRestore(); }
  });
});

describe('conversation history snapshots', () => {
  const chatView = require('../src/taskpane/ui/chat-view.js');
  const { buildConversationHistory } = require('../src/lib/conversation-history.js');

  function setup(actions = makeActions(), extra = {}) {
    document.body.innerHTML = '<div id="chatMessages"></div><div id="welcome"></div>';
    chatView.initChatView();
    chatView.clearChat();
    const view = { ...chatView, renderWelcome: chatView.showWelcome };
    const conv = createConversation({ appState: makeAppState(), view, input: makeInput(),
      actions, log: jest.fn(), getSelectionText: async () => '', ...extra });
    return { conv, view, actions };
  }

  afterEach(() => { chatView.clearChat(); });

  test('uses real same-session records on successive turns without duplicating the latest request', async () => {
    const { conv, actions } = setup();
    await conv.submit('what is the deadline?');
    expect(actions.answerQuestion.mock.calls[0][0].conversationHistory).toEqual([]);
    await conv.submit('what did I just ask?');
    const history = actions.answerQuestion.mock.calls[1][0].conversationHistory;
    expect(history).toEqual([
      { role: 'user', content: 'what is the deadline?' },
      { role: 'assistant', content: expect.stringContaining('the answer') },
    ]);
    expect(history.map((message) => message.content).join('\n')).not.toContain('what did I just ask?');
    expect(actions.answerQuestion.mock.calls[1][1].question).toBe('what did I just ask?');
    expect(chatView.getCurrentSession().messages).toHaveLength(4);
  });

  test('takes a detached snapshot synchronously before selection reading and user insertion', async () => {
    let resolveSelection;
    const { conv, actions } = setup(makeActions(), { getSelectionText: () => new Promise((resolve) => { resolveSelection = resolve; }) });
    chatView.setCurrentSession({ id: 'restored', messages: [
      { role: 'user', text: 'saved question' }, { role: 'assistant', text: 'saved answer',
        proposals: [{ title: 'saved proposal', state: 'pending', items: [{ before: 'old', after: 'new' }] }] },
    ] });
    const running = conv.submit('what is next?');
    const records = chatView.getCurrentSession().messages;
    expect(records).toHaveLength(2);
    records[0].text = 'changed after submit';
    records[1].proposals[0].state = 'applied';
    resolveSelection('');
    await running;
    const history = actions.answerQuestion.mock.calls[0][0].conversationHistory;
    expect(history[0].content).toBe('saved question');
    expect(history[1].content).toContain('pending; proposed only, not applied');
    expect(history[1].content).not.toContain('selected changes were applied');
  });

  test('new chat clears context and restored/switched chats use only their own records', async () => {
    const { conv, actions } = setup();
    await conv.submit('what is first?');
    conv.newChat();
    await conv.submit('what is second?');
    expect(actions.answerQuestion.mock.calls[1][0].conversationHistory).toEqual([]);
    for (const id of ['saved-A', 'saved-B']) {
      chatView.setCurrentSession({ id, messages: [{ role: 'user', text: id }, { role: 'assistant', text: `${id} answer` }] });
      await conv.submit('what is next?');
      expect(actions.answerQuestion.mock.lastCall[0].conversationHistory).toEqual([
        { role: 'user', content: id }, { role: 'assistant', content: `${id} answer` },
      ]);
    }
  });

  test('a direct session switch during selection reading drops the stale submission', async () => {
    let resolveSelection;
    const { conv, actions } = setup(makeActions(), { getSelectionText: () => new Promise((resolve) => { resolveSelection = resolve; }) });
    chatView.setCurrentSession({ id: 'old', messages: [{ role: 'user', text: 'old context' }] });
    const running = conv.submit('what is next?');
    chatView.setCurrentSession({ id: 'new', messages: [{ role: 'user', text: 'new context' }] });
    resolveSelection('');
    await running;
    expect(actions.answerQuestion).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages.map((message) => message.text)).toEqual(['new context']);
  });

  test('a switched-session late answer cannot persist or become subsequent model history', async () => {
    let resolveAnswer;
    const actions = makeActions({ answerQuestion: jest.fn().mockImplementationOnce(() => new Promise((resolve) => { resolveAnswer = resolve; }))
      .mockResolvedValue('fresh answer') });
    const onTurnCommitted = jest.fn();
    const { conv } = setup(actions, { onTurnCommitted });
    chatView.setCurrentSession({ id: 'old', messages: [{ role: 'user', text: 'old context' }] });
    const running = conv.submit('what is old?');
    await Promise.resolve();
    chatView.setCurrentSession({ id: 'new', messages: [{ role: 'user', text: 'new context' }] });
    resolveAnswer('late old answer');
    await running;
    expect(onTurnCommitted).not.toHaveBeenCalled();
    await conv.submit('what is new?');
    expect(actions.answerQuestion.mock.lastCall[0].conversationHistory).toEqual([{ role: 'user', content: 'new context' }]);
  });

  test('cancelled streamed output is never reused, even when the action resolves successfully after abort', async () => {
    let resolveAnswer;
    const actions = makeActions({ answerQuestion: jest.fn().mockImplementationOnce((_deps, options) => {
      options.onToken('partial untrusted answer');
      return new Promise((resolve) => { resolveAnswer = resolve; });
    }).mockResolvedValue('next answer') });
    const { conv } = setup(actions);
    const running = conv.submit('what is this?');
    await Promise.resolve();
    conv.cancel();
    resolveAnswer('partial untrusted answer');
    await running;
    await conv.submit('what happened?');
    const content = actions.answerQuestion.mock.lastCall[0].conversationHistory[1].content;
    expect(content).toContain('Turn cancelled');
    expect(content).not.toContain('partial untrusted answer');
  });

  test('retry closures preserve original history when later turns add more records', async () => {
    let retryDeps;
    const logWithRetry = jest.fn();
    const actions = makeActions({ runDocumentSkill: jest.fn(async (deps) => {
      retryDeps = deps;
      return { results: [], applicationResult: {}, chunks: [], cancelled: false };
    }) });
    const { conv } = setup(actions, { logWithRetry });
    chatView.setCurrentSession({ id: 'retry-session', messages: [{ role: 'user', text: 'original context' }] });
    await conv.submit('polish the document');
    const snapshot = retryDeps.conversationHistory;
    await conv.submit('what happened?');
    await retryDeps.stageRetryProposal({ retryProposal: true, failedCount: 1, results: [{ chunk: { text: 'before' }, amendment: 'after' }],
      chunks: [], retryFailed: jest.fn(), apply: jest.fn(), discard: jest.fn() });
    expect(retryDeps.conversationHistory).toBe(snapshot);
    expect(snapshot).toEqual([{ role: 'user', content: 'original context' }]);
    const retry = jest.fn(() => retryDeps.conversationHistory);
    retryDeps.logWithRetry('retry', 'warning', retry);
    expect(await logWithRetry.mock.lastCall[2]()).toEqual(snapshot);
    chatView.setCurrentSession({ id: 'another', messages: [] });
    await logWithRetry.mock.lastCall[2]();
    expect(retry).toHaveBeenCalledTimes(1);
  });

  test.each([
    ['what is this?', 'answerQuestion'], ['polish the document', 'runDocumentSkill'],
    ['append a conclusion', 'prepareDocumentAppend'], ['make the document bold', 'prepareFormatProposal'],
    ['insert a 3x3 table', 'prepareTableProposal'], ['add an image of a cat', 'prepareIllustrationProposal'],
    ['delete the second image', 'prepareImageToolEdit'], ['delete empty paragraphs', 'prepareEmptyParagraphCleanup'],
    ['/summarize-contract', 'runSummarySkill'], ['增加标题，并深度润色修改', 'planDocumentTasks'],
  ])('passes the same session history into %s via %s', async (instruction, action) => {
    const { conv, actions } = setup();
    chatView.setCurrentSession({ id: 'saved', messages: [{ role: 'user', text: 'prior instruction' }] });
    await conv.submit(instruction);
    expect(actions[action]).toHaveBeenCalled();
    expect(actions[action].mock.calls[0][0].conversationHistory).toEqual([{ role: 'user', content: 'prior instruction' }]);
  });

  test.each([
    ['polish this', { text: 'selected text' }, 'prepareSelectionAmendment'],
    ['summarize this table', { text: 'A B', hasMultiCellTableRegion: true }, 'prepareTableToolEdit'],
  ])('selection context retains conversation history for %s', async (instruction, selection, action) => {
    const { conv, actions } = setup(makeActions(), { getSelectionContent: async () => ({ images: [], ...selection }) });
    chatView.setCurrentSession({ id: 'selection', messages: [{ role: 'user', text: 'prior instruction' }] });
    await conv.submit(instruction);
    expect(actions[action].mock.calls[0][0].conversationHistory).toEqual([{ role: 'user', content: 'prior instruction' }]);
  });

  test('trimming history never trims the current turn and reports its budget decision', async () => {
    const log = jest.fn();
    const appState = makeAppState({ config: { commentGranularity: 0, contextBudgetTokens: 16000 } });
    const { conv, actions } = setup(makeActions(), { appState, log });
    chatView.setCurrentSession({ id: 'large', messages: [
      { role: 'user', text: 'old question' }, { role: 'assistant', text: 'x'.repeat(32001) },
    ] });
    const question = `what is ${'y'.repeat(40000)}?`;
    await conv.submit(question);
    const history = actions.answerQuestion.mock.calls[0][0].conversationHistory;
    expect(history.map((message) => message.role)).toEqual(['user', 'assistant']);
    expect(history[0].content).toBe('old question');
    expect(history[1].content.endsWith(' [trimmed]')).toBe(true);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('Conversation history trimmed'), 'info');
    expect(actions.answerQuestion.mock.calls[0][1].question).toBe(question);
  });

  test('MCP receives the same snapshot option as ordinary actions', async () => {
    const mcp = require('../src/lib/mcp-client.js');
    const loop = require('../src/lib/tool-loop.js');
    const connect = jest.spyOn(mcp, 'connectMcpServer').mockResolvedValue({
      listTools: jest.fn().mockResolvedValue([{ name: 'search', description: 'Search', inputSchema: { type: 'object', properties: {} } }]),
    });
    const run = jest.spyOn(loop, 'runToolLoop').mockResolvedValue({ summary: 'MCP answer' });
    try {
      const { conv } = setup(makeActions(), { appState: makeAppState({ config: {
        backend: 'test', providers: { test: { model: 'model' } }, mcpServers: [{ url: 'https://example.test' }],
      } }) });
      chatView.setCurrentSession({ id: 'mcp-session', messages: [{ role: 'user', text: 'previous MCP request' }] });
      const expected = buildConversationHistory(chatView.getCurrentSession().messages);
      await conv.submit('/mcp follow up');
      expect(run).toHaveBeenCalledWith(expect.objectContaining({ conversationHistory: expected, taskPrompt: 'follow up' }));
    } finally {
      connect.mockRestore();
      run.mockRestore();
    }
  });
});
