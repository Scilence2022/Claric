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

const { routeTurn, createConversation, TURN_TYPE, chunkCitation } = require('../src/taskpane/conversation.js');
const { BUILTIN_SKILLS } = require('../src/taskpane/skills.js');

function makeAppState(overrides = {}) {
  return {
    isProcessing: false,
    isProcessingDoc: false,
    isProcessingSummary: false,
    processDocController: null,
    chatController: null,
    supportsComments: true,
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
  };
  return {
    addUserMessage: jest.fn(),
    createAssistantMessage: jest.fn(() => msg),
    addSystemNote: jest.fn(),
    hideWelcome: jest.fn(),
    renderWelcome: jest.fn(),
    clearChat: jest.fn(),
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
    prepareFormatProposal: jest.fn(async () => ({
      instruction: 'x', scope: 'selection', ops: [{ font: { bold: true } }], model: 'm',
    })),
    applyFormatProposal: jest.fn(async () => ({ applied: 1, warnings: [] })),
    prepareIllustrationProposal: jest.fn(async () => ({
      instruction: 'x',
      svg: '<svg width="10" height="10" viewBox="0 0 10 10"><rect width="10" height="10"/></svg>',
      position: 'end', model: 'm',
    })),
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
    const cardEl = view._msg.attachProposal.mock.calls[0][0];
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
    const cardEl = view._msg.attachProposal.mock.calls[0][0];
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
    const cardEl = view._msg.attachProposal.mock.calls[0][0];
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
    const cardEl = view._msg.attachProposal.mock.calls[0][0];
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
});

describe('createConversation.cancel', () => {
  test('aborts the document and chat controllers', () => {
    const appState = makeAppState({
      processDocController: new AbortController(),
      chatController: new AbortController(),
    });
    const conv = createConversation({
      appState, view: makeView(), input: makeInput(), log: jest.fn(), actions: makeActions(),
    });

    conv.cancel();

    expect(appState.processDocController.signal.aborted).toBe(true);
    expect(appState.chatController.signal.aborted).toBe(true);
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
