/** @jest-environment jsdom */
const { createConversation } = require('../src/taskpane/conversation.js');
const chatView = require('../src/taskpane/ui/chat-view.js');
const { appState } = require('../src/taskpane/app-state.js');

const request = '全文优化格式，例如不正确的字体加粗，多余的空格，表格修改为三线格';

function fixture({ dependencies = true } = {}) {
    const events = [];
    const chunk = { id: 'body', paragraphs: [{ text: 'Example  body.' }] };
    const outcome = {
        staged: true, chunks: [chunk], failedCount: 0, cancelledCount: 0,
        results: [{ status: 'fulfilled', amendment: 'Example body.', chunk, chunkId: chunk.id }],
        discard: jest.fn(async () => {}),
        apply: jest.fn(async () => {
            events.push('apply-body');
            return { amendmentsApplied: 1, commentsInserted: 0 };
        }),
    };
    const actions = {
        planDocumentTasks: jest.fn(async () => ({ tasks: [
            { taskId: 'body', type: 'edit', scope: 'document', instruction: '清理全文非表格正文中多余的空格，仅去除冗余空格，不做内容改写。' },
            { taskId: 'format', type: 'format', scope: 'document', instruction: '纠正全文不合适的加粗', dependsOn: dependencies ? ['body'] : [] },
            { taskId: 'tables', type: 'table_management', scope: 'document', instruction: '表格修改为三线表', dependsOn: dependencies ? ['format'] : [] },
        ] })),
        runDocumentSkill: jest.fn(async () => { events.push('read-body'); return outcome; }),
        prepareFormatProposal: jest.fn(async () => { events.push('read-format'); return { scope: 'document', ops: [{ font: { bold: false } }] }; }),
        discardFormatProposal: jest.fn(async () => {}),
        applyFormatProposal: jest.fn(async () => { events.push('apply-format'); return { appliedRanges: 1, insertedParagraphs: 0 }; }),
        readDocumentTableRegions: jest.fn(async () => { events.push('read-tables'); return [{ tableIndex: 1 }]; }),
        prepareTableToolEdit: jest.fn(async () => ({ tablePatch: {
            cells: [], rowOps: [], merges: [], styleOps: [{ type: 'borders', borders: { all: { type: 'none' } } }],
        }, tableItems: [{ label: 'Three-line table borders' }] })),
        applySelectionAmendment: jest.fn(async () => { events.push('apply-tables'); return {}; }),
        revealTextSnippet: jest.fn(async () => {}),
    };
    const conversation = createConversation({
        appState, actions, view: { ...chatView, renderWelcome: chatView.showWelcome },
        input: { setProcessing: jest.fn(), setValue: jest.fn(), focus: jest.fn() },
        getSelectionText: async () => '', log: jest.fn(),
    });
    return { conversation, actions, outcome, events };
}

function bodyAndTablesFixture({ tablesFirst = false, dependent = false } = {}) {
    const f = fixture();
    const body = { taskId: 'body', type: 'edit', scope: 'document', instruction: '润色全文非表格正文，保持原意。' };
    const tables = { taskId: 'tables', type: 'table_management', scope: 'document', instruction: '将所有现有表格修改为三线格。',
        dependsOn: dependent ? ['body'] : [] };
    f.actions.planDocumentTasks.mockResolvedValue({ tasks: tablesFirst ? [tables, body] : [body, tables] });
    return f;
}

beforeEach(() => {
    jest.useFakeTimers();
    document.body.innerHTML = '<div id="chatMessages"></div><div id="welcome"></div>';
    chatView.initChatView();
    chatView.clearSessionMessages();
    Object.assign(appState, { isProcessing: false, isProcessingDoc: false, isProcessingSummary: false,
        processDocController: null, chatController: null });
    appState.config.autoApplyChanges = true;
    appState.config.trackChangesEnabled = true;
});

afterEach(() => {
    chatView.clearSessionMessages();
    appState.config.autoApplyChanges = false;
    jest.clearAllTimers();
    jest.useRealTimers();
});

test.each([false, true])('independent body and table drafts are both reviewable before Apply (tables first: %s)', async (tablesFirst) => {
    appState.config.autoApplyChanges = false;
    const f = bodyAndTablesFixture({ tablesFirst });
    await f.conversation.submit('全文润色修改，所有表格修改为三线格');
    expect(f.events).toEqual(tablesFirst ? ['read-tables', 'read-body'] : ['read-body', 'read-tables']);
    const message = chatView.getCurrentSession().messages[1];
    expect(message.proposals.map((proposal) => proposal.state)).toEqual(['pending', 'pending']);
    expect(message.status).not.toContain('await application');
    expect(document.body.textContent).not.toContain('task.blocked');
    expect(f.outcome.apply).not.toHaveBeenCalled();
    expect(f.actions.applySelectionAmendment).not.toHaveBeenCalled();

    // Review mode permits applying the independent cards in either order.
    const buttons = [...document.querySelectorAll('.proposal-card button')].filter((button) => button.textContent === 'Apply as tracked changes');
    buttons[1].click();
    await jest.advanceTimersByTimeAsync(10);
    buttons[0].click();
    await jest.advanceTimersByTimeAsync(10);
    expect(message.proposals.map((proposal) => proposal.state)).toEqual(['applied', 'applied']);
    expect(f.outcome.apply).toHaveBeenCalledTimes(1);
    expect(f.actions.applySelectionAmendment).toHaveBeenCalledTimes(1);
});

test('a failed body-generation section does not suppress the independent three-line table draft', async () => {
    appState.config.autoApplyChanges = false;
    const f = bodyAndTablesFixture();
    f.outcome.failedCount = 1;
    const failed = { id: 'failed', paragraphs: [{ text: 'Another synthetic paragraph.' }] };
    f.outcome.chunks.push(failed);
    f.outcome.results.push({ status: 'rejected', chunk: failed, chunkId: failed.id, error: 'Upstream temporarily unavailable' });
    await f.conversation.submit('全文润色修改，所有表格修改为三线格');
    expect(f.actions.prepareTableToolEdit).toHaveBeenCalledTimes(1);
    expect(chatView.getCurrentSession().messages[1].proposals).toHaveLength(2);
    expect(f.events).toEqual(['read-body', 'read-tables']);
});

test('independent drafts auto-apply serially after both preparation steps finish', async () => {
    const f = bodyAndTablesFixture();
    await f.conversation.submit('全文润色修改，所有表格修改为三线格');
    expect(f.events).toEqual(['read-body', 'read-tables']);
    await jest.advanceTimersByTimeAsync(50);
    expect(f.events).toEqual(['read-body', 'read-tables', 'apply-body', 'apply-tables']);
    expect(chatView.getCurrentSession().messages[1].proposals.map((proposal) => proposal.state)).toEqual(['applied', 'applied']);
});

test.each(['error', 'warning', 'paused'])('a body Apply %s leaves the independently drafted table available for review', async (state) => {
    const f = bodyAndTablesFixture();
    if (state === 'error') f.outcome.apply.mockRejectedValue(new Error('Native write failed'));
    else f.outcome.apply.mockResolvedValue(state === 'warning'
        ? { amendmentsApplied: 0, uncertainChunkIds: ['body'], errors: ['Native read-back differs'] }
        : { interrupted: true, amendmentsApplied: 0, appliedChunkIds: [] });
    await f.conversation.submit('全文润色修改，所有表格修改为三线格');
    await jest.advanceTimersByTimeAsync(50);
    expect(f.events).toEqual(['read-body', 'read-tables']);
    expect(f.outcome.apply).toHaveBeenCalledTimes(1);
    expect(f.actions.applySelectionAmendment).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages[1].proposals.map((proposal) => proposal.state)).toEqual([state, 'pending']);
});

test('an explicit dependency between body and table tasks still waits for verified Apply', async () => {
    appState.config.autoApplyChanges = false;
    const f = bodyAndTablesFixture({ dependent: true });
    await f.conversation.submit('先润色正文，应用后再按新正文修改表格为三线格');
    expect(f.actions.readDocumentTableRegions).not.toHaveBeenCalled();
    const button = document.querySelector('.proposal-card .btn-primary');
    button.click();
    await jest.advanceTimersByTimeAsync(50);
    expect(f.events).toEqual(['read-body', 'apply-body', 'read-tables']);
    expect(chatView.getCurrentSession().messages[1].proposals.map((proposal) => proposal.state)).toEqual(['applied', 'pending']);
});

test('a dependency through read-only QA cannot bypass the body Apply boundary', async () => {
    appState.config.autoApplyChanges = false;
    const f = bodyAndTablesFixture();
    const plan = await f.actions.planDocumentTasks();
    plan.tasks.splice(1, 0, { taskId: 'answer', type: 'qa', instruction: 'Explain the planned changes', dependsOn: ['body'] });
    plan.tasks[2].dependsOn = ['answer'];
    f.actions.answerQuestion = jest.fn(async () => 'Synthetic answer');
    await f.conversation.submit('全文润色修改，所有表格修改为三线格');
    expect(f.actions.answerQuestion).toHaveBeenCalledTimes(1);
    expect(f.actions.readDocumentTableRegions).not.toHaveBeenCalled();
    document.querySelector('.proposal-card .btn-primary').click();
    await jest.advanceTimersByTimeAsync(50);
    expect(f.actions.readDocumentTableRegions).toHaveBeenCalledTimes(1);
    expect(f.actions.answerQuestion).toHaveBeenCalledTimes(1);
});

test.each([true, false])('three tasks auto-apply in order with real message/card state (dependencies: %s)', async (dependencies) => {
    const f = fixture({ dependencies });
    await f.conversation.submit(request);
    expect(f.events).toEqual(['read-body']);
    await jest.advanceTimersByTimeAsync(50);
    expect(f.events).toEqual(['read-body', 'apply-body', 'read-format', 'apply-format', 'read-tables', 'apply-tables']);
    const messages = chatView.getCurrentSession().messages;
    expect(messages).toHaveLength(2);
    expect(messages[1].proposals.map((proposal) => proposal.state)).toEqual(['applied', 'applied', 'applied']);
    expect(messages[1].error).toBeNull();
    expect(f.outcome.apply).toHaveBeenCalledTimes(1);
    expect(f.actions.applyFormatProposal).toHaveBeenCalledTimes(1);
    expect(f.actions.applySelectionAmendment).toHaveBeenCalledTimes(1);
});

test('a verified body no-op does not block automatic formatting and table changes', async () => {
    const f = fixture();
    f.actions.runDocumentSkill.mockImplementation(async () => {
        f.events.push('read-body');
        return { status: 'no_op', satisfied: true, summary: 'Verified clean body.', chunks: [], results: [] };
    });
    await f.conversation.submit(request);
    await jest.advanceTimersByTimeAsync(50);
    expect(f.events).toEqual(['read-body', 'read-format', 'apply-format', 'read-tables', 'apply-tables']);
    expect(f.outcome.apply).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages[1].proposals.map((proposal) => proposal.state)).toEqual(['applied', 'applied']);
});

test('turning auto-apply off after the first write leaves the continuation available for review', async () => {
    const f = fixture();
    f.outcome.apply.mockImplementation(async () => {
        f.events.push('apply-body');
        appState.config.autoApplyChanges = false;
        return { amendmentsApplied: 1, commentsInserted: 0 };
    });
    await f.conversation.submit(request);
    await jest.advanceTimersByTimeAsync(50);
    expect(f.events).toEqual(['read-body', 'apply-body', 'read-format']);
    expect(f.actions.applyFormatProposal).not.toHaveBeenCalled();
    expect(f.actions.readDocumentTableRegions).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages[1].proposals.map((proposal) => proposal.state)).toEqual(['applied', 'pending']);
});

test('rejecting the first proposal before its deferred write stops the chain', async () => {
    const f = fixture();
    await f.conversation.submit(request);
    const reject = [...document.querySelectorAll('.proposal-card button')].find((button) => button.textContent === 'Reject');
    reject.click();
    await jest.advanceTimersByTimeAsync(50);
    expect(f.outcome.apply).not.toHaveBeenCalled();
    expect(f.actions.prepareFormatProposal).not.toHaveBeenCalled();
    expect(f.actions.readDocumentTableRegions).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages[1].proposals[0].state).toBe('rejected');
});

test('a failed automatic Apply does not retry a native write or start downstream tasks', async () => {
    const f = fixture();
    f.outcome.apply.mockRejectedValue(new Error('Native write refused'));
    await f.conversation.submit(request);
    await jest.advanceTimersByTimeAsync(50);
    expect(f.outcome.apply).toHaveBeenCalledTimes(1);
    expect(f.actions.prepareFormatProposal).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages[1].proposals[0].state).toBe('error');
});

test('an unverified item overrides an optimistic aggregate count and stops dependent automatic edits', async () => {
    const f = fixture();
    f.outcome.apply.mockImplementation(async (_ids, { onChunkApplied }) => {
        onChunkApplied('body', { applied: false, uncertain: true });
        return { amendmentsApplied: 1, commentsInserted: 0 };
    });
    await f.conversation.submit(request);
    await jest.advanceTimersByTimeAsync(50);
    expect(f.outcome.apply).toHaveBeenCalledTimes(1);
    expect(f.actions.prepareFormatProposal).not.toHaveBeenCalled();
    expect(f.actions.readDocumentTableRegions).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages[1].proposals[0].state).toBe('warning');
    expect(chatView.getCurrentSession().messages[1].status).not.toMatch(/^Done:/);
    expect(document.querySelector('.proposal-applied')).toBeNull();
});

test('Stop during automatic Apply pauses its card and does not resume or retry the chain', async () => {
    const f = fixture();
    f.outcome.apply.mockImplementation(async (_ids, { signal }) => new Promise((resolve) => {
        signal.addEventListener('abort', () => resolve({ interrupted: true, appliedChunkIds: [] }), { once: true });
    }));
    await f.conversation.submit(request);
    await jest.advanceTimersByTimeAsync(0);
    expect(f.outcome.apply).toHaveBeenCalledTimes(1);
    f.conversation.cancel();
    await jest.advanceTimersByTimeAsync(50);
    expect(f.outcome.apply).toHaveBeenCalledTimes(1);
    expect(f.actions.prepareFormatProposal).not.toHaveBeenCalled();
    expect(document.querySelector('.proposal-paused')).not.toBeNull();
});

test('Stop during later independent model work prevents deferred automatic writes from the cancelled turn', async () => {
    const f = fixture();
    const plan = await f.actions.planDocumentTasks();
    plan.tasks[2] = { taskId: 'answer', type: 'qa', instruction: 'Explain the cleanup', dependsOn: [] };
    f.actions.planDocumentTasks.mockResolvedValue(plan);
    f.actions.answerQuestion = jest.fn(async (_deps, { signal }) => new Promise((_resolve, reject) => {
        signal.addEventListener('abort', () => reject(new DOMException('Cancelled', 'AbortError')), { once: true });
    }));
    const submitted = f.conversation.submit(request);
    for (let i = 0; i < 30 && !f.actions.answerQuestion.mock.calls.length; i++) await Promise.resolve();
    expect(f.actions.answerQuestion).toHaveBeenCalledTimes(1);
    f.conversation.cancel();
    await submitted;
    await jest.advanceTimersByTimeAsync(50);
    expect(f.outcome.apply).not.toHaveBeenCalled();
    expect(f.actions.applyFormatProposal).not.toHaveBeenCalled();
});

test('changing the chat before a deferred automatic Apply prevents old-document writes', async () => {
    const f = fixture();
    await f.conversation.submit(request);
    f.conversation.newChat();
    await jest.advanceTimersByTimeAsync(50);
    expect(f.outcome.apply).not.toHaveBeenCalled();
    expect(f.actions.prepareFormatProposal).not.toHaveBeenCalled();
    expect(chatView.getCurrentSession().messages).toEqual([]);
});
