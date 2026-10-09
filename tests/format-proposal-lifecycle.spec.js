/** @jest-environment jsdom */
const { createConversation } = require('../src/taskpane/conversation.js');
const view = require('../src/taskpane/ui/chat-view.js');
const { appState } = require('../src/taskpane/app-state.js');

beforeEach(() => {
    document.body.innerHTML = '<div id="chatMessages"></div><div id="welcome"></div>';
    view.initChatView(); view.clearSessionMessages();
    appState.config.autoApplyChanges = true;
    appState.config.trackChangesEnabled = true;
    appState.isProcessing = false; appState.isProcessingDoc = false; appState.isProcessingSummary = false;
    appState.chatController = null; appState.processDocController = null;
});
afterEach(() => { appState.config.autoApplyChanges = false; view.clearChat(); });

function setup(actions, selectionText = 'Example heading\n\nExample body') {
    const cards = [];
    const liveView = { ...view, createAssistantMessage(...args) {
        const msg = view.createAssistantMessage(...args);
        const attach = msg.attachProposal;
        msg.attachProposal = (card, meta) => { cards.push(card); attach(card, meta); };
        return msg;
    } };
    const conversation = createConversation({ appState, view: liveView, log: jest.fn(), actions,
        input: { setProcessing: jest.fn(), setValue: jest.fn(), focus: jest.fn() },
        getSelectionText: async () => selectionText });
    return { conversation, cards };
}

test('automatic apply preserves the first native error and a consumed formatting card cannot retry', async () => {
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ font: { bold: false } }], anchor: { bookmark: '_test' } })),
        applyFormatProposal: jest.fn(async () => { throw new Error('The anchored formatting baseline changed. Draft a new proposal.'); }),
        discardFormatProposal: jest.fn(async () => {}),
    };
    const { conversation, cards } = setup(actions);
    await conversation.submit('调整选择部分的格式，包括多余的空行，不正确的粗体格式等');
    await new Promise((resolve) => setTimeout(resolve, 0));
    const card = cards[0];
    expect(actions.applyFormatProposal).toHaveBeenCalledTimes(1);
    expect(actions.discardFormatProposal).toHaveBeenCalledTimes(1);
    const button = card.el.querySelector('.btn-primary');
    expect(button.disabled).toBe(true);
    button.click(); await card.applyAll();
    card.markError('This proposal has been discarded.');
    card.markApplied('late success');
    expect(actions.applyFormatProposal).toHaveBeenCalledTimes(1);
    expect(card.el.textContent).toContain('The anchored formatting baseline changed');
    expect(card.el.textContent).not.toContain('This proposal has been discarded');
    const record = view.getCurrentSession().messages.find((message) => message.role === 'assistant');
    expect(record.proposals[0]).toMatchObject({ state: 'error', detail: expect.stringContaining('baseline changed') });
    view.setCurrentSession(view.getCurrentSession());
    expect(document.querySelector('.proposal-card button')).toBeNull();
});

test('cleanup-only success is reported as applied, with one reviewable deletion item', async () => {
    appState.config.autoApplyChanges = false;
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ cleanup: { emptyParagraphs: true, emptyCount: 2 } }], anchor: { bookmark: '_test' } })),
        applyFormatProposal: jest.fn(async () => ({ applied: true, appliedRanges: 0, insertedParagraphs: 0, deletedParagraphs: 2 })),
        discardFormatProposal: jest.fn(async () => {}),
    };
    const { conversation, cards } = setup(actions);
    await conversation.submit('删除选区多余空行');
    expect(actions.prepareFormatProposal.mock.calls[0][1]).toMatchObject({ scope: 'selection', cleanupOnly: true, cleanupRequested: true });
    expect(actions.applyFormatProposal).not.toHaveBeenCalled();
    expect(cards[0].el.textContent).toContain('Delete 2 verified empty paragraph');
    await cards[0].applyAll();
    expect(cards[0].el.classList.contains('proposal-applied')).toBe(true);
    expect(view.getCurrentSession().messages.find((m) => m.role === 'assistant').proposals[0].state).toBe('applied');
});

test('successful formatting does not claim blank-line cleanup succeeded when all blanks were protected', async () => {
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ font: { bold: false } }], anchor: { bookmark: '_test' },
            cleanupSummary: { candidates: 2, verified: 0, preserved: 2, unverifiable: 1 } })),
        applyFormatProposal: jest.fn(async () => ({ applied: true, appliedRanges: 1, insertedParagraphs: 0, deletedParagraphs: 0 })),
        discardFormatProposal: jest.fn(async () => {}),
    };
    const { conversation, cards } = setup(actions);
    await conversation.submit('整理选择部分的格式，多余的空行，不合适的字体加粗等');
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(cards[0].el.textContent).toContain('2 empty paragraph(s) preserved');
    expect(cards[0].el.textContent).toContain('Formatting applied; 2 empty paragraph(s) could not be safely removed');
    expect(view.getCurrentSession().messages.find((m) => m.role === 'assistant').proposals[0].state).toBe('warning');
});

test('cleanup without safe targets explains preservation instead of claiming no blanks exist', async () => {
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [], anchor: { bookmark: '_test' },
            cleanupSummary: { candidates: 2, verified: 0, preserved: 2, unverifiable: 1 } })),
        discardFormatProposal: jest.fn(async () => {}),
        applyFormatProposal: jest.fn(),
    };
    const { conversation } = setup(actions);
    await conversation.submit('删除选区多余空行');
    expect(document.getElementById('chatMessages').textContent).toContain('2 empty paragraph(s) could not be safely removed');
    expect(actions.applyFormatProposal).not.toHaveBeenCalled();
});

test('verified cleanup reports removed and structurally preserved paragraphs as applied', async () => {
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ font: { bold: false } }, { cleanup: { emptyParagraphs: true, emptyCount: 3 } }],
            anchor: { bookmark: '_test' }, cleanupSummary: { candidates: 5, verified: 3, preserved: 2, unverifiable: 0 } })),
        applyFormatProposal: jest.fn(async () => ({ applied: true, appliedRanges: 1, insertedParagraphs: 0, deletedParagraphs: 3 })),
        discardFormatProposal: jest.fn(async () => {}),
    };
    const { conversation, cards } = setup(actions);
    await conversation.submit('整理选择部分的格式，多余的空行，不合适的字体加粗等');
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(cards[0].el.textContent).toContain('3 empty paragraph(s) removed; 2 protected empty paragraph(s) preserved');
    expect(view.getCurrentSession().messages.find((m) => m.role === 'assistant').proposals[0].state).toBe('applied');
});

test('a verified formatting no-op is persisted as satisfied rather than mistaken for zero matched targets', async () => {
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ paragraphRole: 'body', paragraph: { alignment: 'justified' } }],
            anchor: { bookmark: '_test' }, targetSummary: { verifiedParagraphs: 3, excludedParagraphs: 2, uncertainParagraphs: 0 } })),
        applyFormatProposal: jest.fn(async () => ({ applied: false, appliedRanges: 0, insertedParagraphs: 0,
            deletedParagraphs: 0, alreadySatisfied: true, verifiedParagraphs: 3, noopParagraphs: 3 })),
        discardFormatProposal: jest.fn(async () => {}),
    };
    const { conversation, cards } = setup(actions);
    await conversation.submit('整理选择部分的格式');
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(cards[0].el.textContent).toContain('3 verified paragraph(s)');
    expect(cards[0].el.textContent).toContain('requested formatting was already satisfied');
    expect(cards[0].el.textContent).not.toContain('no formatting targets matched');
    expect(cards[0].el.classList.contains('proposal-applied')).toBe(true);
    const proposal = view.getCurrentSession().messages.find((message) => message.role === 'assistant').proposals[0];
    expect(proposal).toMatchObject({ state: 'applied', detail: expect.stringContaining('already satisfied') });
    await cards[0].applyAll();
    expect(actions.applyFormatProposal).toHaveBeenCalledTimes(1);
    view.setCurrentSession(view.getCurrentSession());
    expect(document.getElementById('chatMessages').textContent).toContain('already satisfied');
    expect(document.querySelector('.proposal-card button')).toBeNull();
});

test('zero targets never becomes an already-satisfied formatting claim without native verification', async () => {
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ paragraphStyle: 'normal', paragraph: { alignment: 'justified' } }],
            anchor: { bookmark: '_test' } })),
        applyFormatProposal: jest.fn(async () => ({ applied: false, appliedRanges: 0, insertedParagraphs: 0,
            deletedParagraphs: 0, alreadySatisfied: true, verifiedParagraphs: 0, noopParagraphs: 0 })),
        discardFormatProposal: jest.fn(async () => {}),
    };
    const { conversation, cards } = setup(actions);
    await conversation.submit('整理选择部分的格式');
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(cards[0].el.textContent).toContain('no formatting targets matched');
    expect(cards[0].el.textContent).not.toContain('requested formatting was already satisfied');
    const proposal = view.getCurrentSession().messages.find((message) => message.role === 'assistant').proposals[0];
    expect(proposal).toMatchObject({ state: 'warning', detail: expect.stringContaining('Nothing applied') });
    await cards[0].applyAll();
    expect(actions.applyFormatProposal).toHaveBeenCalledTimes(1);
});

test.each([false, true])('uncertain body paragraphs remain a warning even when verified targets are satisfied: no-op=%s', async (alreadySatisfied) => {
    const actions = {
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ paragraphRole: 'body', paragraph: { alignment: 'justified' } }],
            anchor: { bookmark: '_test' }, targetSummary: { verifiedParagraphs: 3, excludedParagraphs: 5, uncertainParagraphs: 2 } })),
        applyFormatProposal: jest.fn(async () => ({ applied: !alreadySatisfied, appliedRanges: alreadySatisfied ? 0 : 3,
            insertedParagraphs: 0, deletedParagraphs: 0, alreadySatisfied, verifiedParagraphs: 3, noopParagraphs: alreadySatisfied ? 3 : 0 })),
        discardFormatProposal: jest.fn(async () => {}),
    };
    const { conversation, cards } = setup(actions);
    await conversation.submit('整理选择部分的格式');
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(cards[0].el.textContent).toContain('3 verified paragraph(s)');
    expect(cards[0].el.textContent).toContain('2 uncertain paragraph(s) preserved');
    expect(cards[0].el.textContent).toContain('3 paragraph(s) verified; 2 uncertain paragraph(s) preserved for review');
    expect(cards[0].el.textContent).not.toContain('requested formatting was already satisfied');
    const proposal = view.getCurrentSession().messages.find((message) => message.role === 'assistant').proposals[0];
    expect(proposal).toMatchObject({ state: 'warning', detail: expect.stringContaining('preserved for review') });
    view.setCurrentSession(view.getCurrentSession());
    expect(document.getElementById('chatMessages').textContent).toContain('preserved for review');
    expect(document.querySelector('.proposal-card button')).toBeNull();
});

test('a composite planner-expanded format task preserves the original request at native preparation', async () => {
    appState.config.autoApplyChanges = false;
    const request = '正文修改为两端对齐，然后创建表格';
    const instruction = '将文档中的正文段落设置为两端对齐（Justified）。仅调整正文，不修改标题、表格或图片。';
    const actions = {
        planDocumentTasks: jest.fn(async () => ({ tasks: [
            { taskId: 'format', type: 'format', scope: 'document', instruction, dependsOn: [] },
            { taskId: 'table', type: 'table', scope: 'document', instruction: 'Create a table', dependsOn: ['format'] },
        ] })),
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ paragraphRole: 'body', paragraph: { alignment: 'justified' } }],
            anchor: { kind: 'document-body' }, targetSummary: { verifiedParagraphs: 1, excludedParagraphs: 0, uncertainParagraphs: 0 } })),
        applyFormatProposal: jest.fn(), discardFormatProposal: jest.fn(async () => {}),
        prepareTableProposal: jest.fn(),
    };
    const { conversation, cards } = setup(actions, '');
    await conversation.submit(request);
    expect(actions.planDocumentTasks).toHaveBeenCalledTimes(1);
    expect(actions.prepareFormatProposal).toHaveBeenCalledWith(expect.anything(), expect.objectContaining({
        instruction, originalInstruction: request, scope: 'document', selectionText: '',
    }));
    expect(actions.prepareTableProposal).not.toHaveBeenCalled();
    expect(cards).toHaveLength(1);
    expect(cards[0].el.textContent).toContain('1 verified paragraph(s)');
});

test.each(['verified-no-op', 'zero-targets', 'uncertain'])('dependent native tasks resume only after a fully verified formatting result: %s', async (mode) => {
    appState.config.autoApplyChanges = false;
    appState.supportsTables = true;
    const verifiedParagraphs = mode === 'zero-targets' ? 0 : 2;
    const uncertainParagraphs = mode === 'uncertain' ? 1 : 0;
    const actions = {
        planDocumentTasks: jest.fn(async () => ({ tasks: [
            { taskId: 'format', type: 'format', scope: 'document', instruction: '将正文段落设置为两端对齐。', dependsOn: [] },
            { taskId: 'table', type: 'table', scope: 'document', instruction: 'Create a table', dependsOn: ['format'] },
        ] })),
        prepareFormatProposal: jest.fn(async () => ({ ops: [{ paragraphRole: 'body', paragraph: { alignment: 'justified' } }],
            anchor: { kind: 'document-body' }, targetSummary: { verifiedParagraphs, excludedParagraphs: uncertainParagraphs, uncertainParagraphs } })),
        applyFormatProposal: jest.fn(async () => ({ applied: false, appliedRanges: 0, insertedParagraphs: 0,
            deletedParagraphs: 0, alreadySatisfied: true, verifiedParagraphs, noopParagraphs: verifiedParagraphs })),
        discardFormatProposal: jest.fn(async () => {}),
        prepareTableProposal: jest.fn(async () => ({ spec: { rows: [['A', 'B']], position: 'end', headerRowCount: 1, style: 'tableGrid' } })),
    };
    const { conversation, cards } = setup(actions, '');
    await conversation.submit('正文修改为两端对齐，然后创建表格');
    expect(actions.prepareTableProposal).not.toHaveBeenCalled();
    await cards[0].applyAll();
    await new Promise((resolve) => setTimeout(resolve, 20));
    if (mode === 'verified-no-op') {
        expect(actions.prepareTableProposal).toHaveBeenCalledTimes(1);
        expect(cards).toHaveLength(2);
        expect(cards[0].el.classList.contains('proposal-applied')).toBe(true);
    } else {
        expect(actions.prepareTableProposal).not.toHaveBeenCalled();
        expect(cards).toHaveLength(1);
        const proposal = view.getCurrentSession().messages.find((message) => message.role === 'assistant').proposals[0];
        expect(proposal.state).toBe('warning');
    }
    expect(actions.prepareFormatProposal).toHaveBeenCalledTimes(1);
    expect(actions.applyFormatProposal).toHaveBeenCalledTimes(1);
});
