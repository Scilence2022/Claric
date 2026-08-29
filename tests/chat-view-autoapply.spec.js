/** @jest-environment jsdom */

/**
 * Auto-apply pipeline: a card staged while config.autoApplyChanges is on
 * fires its applyAll when the turn finalizes — never mid-turn (the write
 * must not race the turn's own busy-flag teardown).
 */

const {
    initChatView,
    createAssistantMessage,
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
        appState.config.autoApplyChanges = false;
    });

    afterEach(() => {
        appState.config.autoApplyChanges = false;
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
