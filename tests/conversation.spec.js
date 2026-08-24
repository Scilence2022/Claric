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

  test('empty input returns null', () => {
    expect(routeTurn('   ', { hasSelection: false, skills: BUILTIN_SKILLS })).toBeNull();
    expect(routeTurn('', { hasSelection: true, skills: BUILTIN_SKILLS })).toBeNull();
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
      actions, getSelectionState: async () => true,
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
      actions, getSelectionState: async () => false,
    });

    await conv.submit('what is this document about?');

    expect(actions.answerQuestion).toHaveBeenCalledTimes(1);
    expect(actions.answerQuestion.mock.calls[0][1].question).toBe('what is this document about?');
    expect(view._msg.setText).toHaveBeenCalledWith('the answer');
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => true,
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => false,
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
      actions, getSelectionState: async () => false,
    });

    await conv.submit('/summarize-contract');

    expect(view._msg.markError).toHaveBeenCalledWith(expect.stringContaining('boom'));
    expect(appState.isProcessingSummary).toBe(false);
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
