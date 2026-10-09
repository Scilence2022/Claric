/** @jest-environment jsdom */
const { createConversation } = require('../src/taskpane/conversation.js');

const instruction = '清理全文非表格正文中多余的空格；仅去除冗余空格，保留有意义的空格、词语和段落结构，不做内容改写。';

function fixture({ original = 'Example body.', amendment = original, taskInstruction = instruction, resultsCount = 1, chunksCount = 1 } = {}) {
  const chunk = { id: 'body', paragraphs: original.split('\n').map((text) => ({ text })) };
  const outcome = { staged: true, results: Array.from({ length: resultsCount }, () => ({ status: 'fulfilled', amendment, chunk, chunkId: 'body' })),
    chunks: Array.from({ length: chunksCount }, () => chunk), failedCount: 0, cancelledCount: 0, discard: jest.fn(async () => {}),
    apply: jest.fn(async () => ({ amendmentsApplied: 1, commentsInserted: 0 })) };
  const msg = Object.fromEntries(['setStatus', 'setText', 'appendText', 'appendLogLine', 'collapseLog', 'appendModelToken',
    'collapseModelOutput', 'showProgress', 'hideProgress', 'attachProposal', 'addCitationPills', 'markError', 'finalizeForHistory'].map((name) => [name, jest.fn()]));
  const view = { createAssistantMessage: () => msg, addUserMessage: jest.fn(), hideWelcome: jest.fn(),
    getCurrentSession: () => ({ id: 's', messages: [] }) };
  const actions = { planDocumentTasks: jest.fn(async () => ({ tasks: [{ taskId: 'body', type: 'edit', scope: 'document', instruction: taskInstruction }] })),
    runDocumentSkill: jest.fn(async () => outcome), readSelectionSnippet: async () => '', revealTextSnippet: jest.fn() };
  const appState = { isProcessing: false, isProcessingDoc: false, isProcessingSummary: false,
    config: { commentGranularity: 0 }, promptManager: { getPrompts: () => [], getActivePrompt: () => null } };
  const input = { setProcessing: jest.fn(), setValue: jest.fn(), focus: jest.fn() };
  const conv = createConversation({ appState, view, input, log: jest.fn(), actions, getSelectionText: async () => '' });
  return { conv, msg, actions, outcome, events: () => msg.appendModelToken.mock.calls.map((call) => call[2]).join('') };
}

async function submit(f) {
  await f.conv.submit('全文优化格式，例如不正确的字体加粗，多余的空格，表格修改为三线格');
}

test('fully processed clean body text yields an explicitly verified no-op', async () => {
  const f = fixture(); await submit(f);
  expect(f.actions.runDocumentSkill.mock.calls[0][1].whitespaceOnly).toBe(true);
  expect(f.msg.attachProposal).not.toHaveBeenCalled();
  expect(f.events()).toContain('task.no_op');
  expect(f.events()).not.toContain('task.failed');
  expect(f.outcome.discard).toHaveBeenCalledWith({ unchanged: true });
});

test('an unchanged model echo cannot satisfy a request with remaining source candidates', async () => {
  const f = fixture({ original: 'Example  body.' }); await submit(f);
  expect(f.events()).toContain('task.failed');
  expect(f.events()).toContain('Potential redundant-space candidates remain');
  expect(f.msg.attachProposal).not.toHaveBeenCalled();
});

test('potentially meaningful ASCII indentation and equation alignment require a model decision', async () => {
  const f = fixture({ original: '  x  =  y' }); await submit(f);
  expect(f.events()).toContain('task.failed');
  expect(f.msg.attachProposal).not.toHaveBeenCalled();
  expect(f.outcome.apply).not.toHaveBeenCalled();
});

test('missing processing coverage cannot be a verified no-op', async () => {
  const f = fixture({ chunksCount: 2 }); await submit(f);
  expect(f.events()).toContain('task.failed');
  expect(f.events()).toContain('not verified for every document section');
});

test('generic polishing echoes are never promoted to a verified space-cleanup no-op', async () => {
  const f = fixture({ taskInstruction: 'Polish the full body' }); await submit(f);
  expect(f.actions.runDocumentSkill.mock.calls[0][1].whitespaceOnly).toBeUndefined();
  expect(f.events()).toContain('task.failed');
});

test('boundary-space deletion creates an applicable proposal', async () => {
  const f = fixture({ original: '  Example body.  ', amendment: 'Example body.' }); await submit(f);
  expect(f.msg.attachProposal).toHaveBeenCalledTimes(1);
  expect(f.outcome.results[0].whitespaceOnly).toBe(true);
  await f.msg.attachProposal.mock.calls[0][0].applyAll();
  expect(f.outcome.apply).toHaveBeenCalled();
});

test.each(['Changed body.', 'Examplebody.'])('non-space rewriting and word joining are rejected before staging: %s', async (amendment) => {
  const f = fixture({ original: 'Example body.', amendment }); await submit(f);
  expect(f.msg.attachProposal).not.toHaveBeenCalled();
  expect(f.outcome.apply).not.toHaveBeenCalled();
  expect(f.events()).toContain('task.failed');
  expect(f.outcome.results[0].status).toBe('rejected');
});
