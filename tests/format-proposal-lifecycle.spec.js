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
