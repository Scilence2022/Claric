/** @jest-environment jsdom */

/**
 * Auto-apply pipeline: a card staged while config.autoApplyChanges is on
 * fires its applyAll when the turn finalizes — never mid-turn (the write
 * must not race the turn's own busy-flag teardown).
 */

const {
    initChatView,
    createAssistantMessage,
    clearChat,
    clearSessionMessages,
    setCurrentSession,
    getCurrentSession,
} = require('../src/taskpane/ui/chat-view.js');
const { appState } = require('../src/taskpane/app-state.js');

function setupDom() {
    document.body.innerHTML = '<div id="chatMessages"></div><div id="welcome"></div>';
    initChatView();
}

function makeCard() {
    const el = document.createElement('div');
    return {
        el,
        applyAll: jest.fn(),
        markApplied: jest.fn(),
    };
}

describe('auto-apply pipeline', () => {
    beforeEach(() => {
        setupDom();
        clearSessionMessages();
        appState.config.trackChangesEnabled = true;
        appState.config.autoApplyChanges = false;
    });

    afterEach(() => {
        appState.config.autoApplyChanges = false;
    });

    const transitions = {
        clearChat: () => clearChat(),
        resetWithoutRemovingDOM: () => clearSessionMessages(),
        loadSession: () => setCurrentSession({ id: 'another', messages: [] }),
        reloadSameSession: () => setCurrentSession(getCurrentSession()),
    };

    test.each(Object.keys(transitions))('%s cancels a deferred queue before its first write', async (transition) => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const card = makeCard();
        msg.attachProposal(card, { title: 'T', state: 'pending', items: [] });
        msg.finalizeForHistory();
        transitions[transition]();
        await new Promise((r) => setTimeout(r, 0));
        expect(card.applyAll).not.toHaveBeenCalled();
    });

    test.each(Object.keys(transitions))('%s during an awaited apply stops remaining cards', async (transition) => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        let release;
        const first = makeCard();
        first.applyAll.mockImplementation(() => new Promise((resolve) => { release = resolve; }));
        const next = makeCard();
        msg.attachProposal(first, { title: 'First', state: 'pending', items: [] });
        msg.attachProposal(next, { title: 'Next', state: 'pending', items: [] });
        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(first.applyAll).toHaveBeenCalledTimes(1);
        transitions[transition]();
        release();
        await new Promise((r) => setTimeout(r, 0));
        expect(next.applyAll).not.toHaveBeenCalled();
    });

    test('detached DOM alone does not change the session identity', async () => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const card = makeCard();
        msg.attachProposal(card, { title: 'T', state: 'pending', items: [] });
        document.getElementById('chatMessages').remove();
        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(card.applyAll).toHaveBeenCalledTimes(1);
    });

    test('late finalization cannot write old messages into a new session', async () => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const card = makeCard();
        msg.attachProposal(card, { title: 'T', state: 'pending', items: [] });
        clearChat();
        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(card.applyAll).not.toHaveBeenCalled();
        expect(getCurrentSession().messages).toEqual([]);
    });

    test.each(['markApplied', 'markRejected', 'markWarning'])('history preserves %s as terminal after late retry callbacks', async (method) => {
        const { createProposalCard } = require('../src/taskpane/ui/proposal-card.js');
        const msg = createAssistantMessage();
        const onApply = jest.fn();
        const card = createProposalCard({ title: 'T', onApply });
        const meta = { title: 'T', state: 'pending', items: [] };
        msg.attachProposal(card, meta);
        msg.finalizeForHistory();
        card[method]();
        const terminalState = meta.state;
        card.markError('late failure');
        card.setPaused('late pause');
        await card.applyAll();
        expect(meta.state).toBe(terminalState);
        expect(getCurrentSession().messages[0].proposals[0].state).toBe(terminalState);
        setCurrentSession(getCurrentSession());
        expect(document.querySelector('.proposal-card button')).toBeNull();
        expect(onApply).not.toHaveBeenCalled();
    });

    test('with auto-apply OFF nothing is scheduled', () => {
        const msg = createAssistantMessage();
        const card = makeCard();
        msg.attachProposal(card, { title: 'T', state: 'pending', items: [] });
        msg.finalizeForHistory();
        expect(card.applyAll).not.toHaveBeenCalled();
    });

    test('with auto-apply ON the pending card applies at finalize, not before', async () => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const card = makeCard();
        msg.attachProposal(card, { title: 'T', state: 'pending', items: [] });

        // Armed but not fired mid-turn.
        expect(card.applyAll).not.toHaveBeenCalled();

        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(card.applyAll).toHaveBeenCalledTimes(1);
    });

    test('with auto-apply ON every pending card applies (multi-card turns)', async () => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const first = makeCard();
        const second = makeCard();
        // Multi-pipeline runs stage one proposal card per pipeline on the
        // same assistant message.
        msg.attachProposal(first, { title: 'T1', state: 'pending', items: [] });
        msg.attachProposal(second, { title: 'T2', state: 'pending', items: [] });

        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(first.applyAll).toHaveBeenCalledTimes(1);
        expect(second.applyAll).toHaveBeenCalledTimes(1);
    });

    test('cards apply sequentially — the next waits for the in-flight one', async () => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        let releaseFirst;
        const first = makeCard();
        // The card's real runApply owns a cross-card mutex while in flight,
        // so the queue must await each apply before starting the next.
        first.applyAll = jest.fn(() => new Promise((resolve) => { releaseFirst = resolve; }));
        const second = makeCard();
        msg.attachProposal(first, { title: 'T1', state: 'pending', items: [] });
        msg.attachProposal(second, { title: 'T2', state: 'pending', items: [] });

        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(first.applyAll).toHaveBeenCalledTimes(1);
        expect(second.applyAll).not.toHaveBeenCalled();

        releaseFirst();
        await new Promise((r) => setTimeout(r, 0));
        expect(second.applyAll).toHaveBeenCalledTimes(1);
    });

    test.each(['autoApplyChanges', 'trackChangesEnabled'])('turning off %s stops queued cards', async (setting) => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const first = makeCard();
        const second = makeCard();
        first.applyAll.mockImplementation(async () => { appState.config[setting] = false; });
        msg.attachProposal(first, { title: 'T1', state: 'pending', items: [] });
        msg.attachProposal(second, { title: 'T2', state: 'pending', items: [] });
        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(first.applyAll).toHaveBeenCalledTimes(1);
        expect(second.applyAll).not.toHaveBeenCalled();
    });

    test.each(['error', 'tracking off', 'disabled before timer'])('%s prevents automatic writes', async (reason) => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const card = makeCard();
        msg.attachProposal(card, { title: 'T', state: 'pending', items: [] });
        if (reason === 'error') msg.markError('Failed turn');
        if (reason === 'tracking off') appState.config.trackChangesEnabled = false;
        msg.finalizeForHistory();
        if (reason === 'disabled before timer') appState.config.autoApplyChanges = false;
        await new Promise((r) => setTimeout(r, 0));
        expect(card.applyAll).not.toHaveBeenCalled();
    });

    test('a queued card that is rejected before the timer does not apply', async () => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const card = makeCard();
        const meta = { title: 'T', state: 'pending', items: [] };
        msg.attachProposal(card, meta);
        msg.finalizeForHistory();
        meta.state = 'rejected';
        await new Promise((r) => setTimeout(r, 0));
        expect(card.applyAll).not.toHaveBeenCalled();
    });

    test('non-pending metadata is never auto-applied', async () => {
        appState.config.autoApplyChanges = true;
        const msg = createAssistantMessage();
        const card = makeCard();
        msg.attachProposal(card, { title: 'T', state: 'applied', items: [] });
        msg.finalizeForHistory();
        await new Promise((r) => setTimeout(r, 0));
        expect(card.applyAll).not.toHaveBeenCalled();
    });
});
