/** @jest-environment jsdom */

/**
 * History view tests.
 *
 * Covers:
 *   - openHistory renders one row per saved session, newest first
 *   - row click fires onLoadSession with the full session payload
 *   - delete button fires onDeleteSession and removes the row
 *   - empty state shows when no sessions exist
 *   - close button hides the overlay
 */

const {
    initHistoryView,
    openHistory,
    closeHistory,
    isHistoryOpen,
} = require('../src/taskpane/ui/history-view.js');

const sessions = require('../src/taskpane/sessions.js');

// localStorage mock (node test environment has no DOM)
const localStorageMock = (() => {
    let store = {};
    return {
        getItem: (key) => (key in store ? store[key] : null),
        setItem: (key, value) => { store[key] = String(value); },
        removeItem: (key) => { delete store[key]; },
        clear: () => { store = {}; },
        get length() { return Object.keys(store).length; },
        key: (i) => Object.keys(store)[i] || null,
    };
})();
global.localStorage = localStorageMock;

function setupDom() {
    document.body.innerHTML = `
        <button id="historyBtn"></button>
        <div id="historyOverlay" hidden>
            <aside class="history-panel">
                <div class="history-header">
                    <button id="historyNewChatBtn">New chat</button>
                    <button id="historyCloseBtn">&times;</button>
                </div>
                <div id="historyBody"></div>
            </aside>
        </div>
    `;
}

beforeEach(() => {
    setupDom();
    sessions.clearAllSessions();
});

describe('openHistory', () => {
    test('renders one row per saved session, newest first', () => {
        sessions.saveSession([{ id: 'm-0', role: 'user', text: 'first', ts: 'x', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] }]);
        sessions.saveSession([{ id: 'm-0', role: 'user', text: 'second', ts: 'x', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] }]);
        sessions.saveSession([{ id: 'm-0', role: 'user', text: 'third', ts: 'x', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] }]);

        const list = openHistory();

        expect(list).toHaveLength(3);
        const titles = list.map((s) => s.title);
        expect(titles[0]).toBe('third');
        expect(titles[2]).toBe('first');
        // DOM: rows are present
        expect(document.querySelectorAll('.history-item')).toHaveLength(3);
        expect(isHistoryOpen()).toBe(true);
        // overlay is visible
        expect(document.getElementById('historyOverlay').hasAttribute('hidden')).toBe(false);
    });

    test('renders the empty state when there are no sessions', () => {
        const list = openHistory();
        expect(list).toEqual([]);
        expect(document.querySelector('.history-empty')).not.toBeNull();
        expect(document.querySelector('.history-empty').textContent).toMatch(/no saved conversations/i);
        expect(isHistoryOpen()).toBe(true);
    });

    test('clicking a row fires onLoadSession with the full session payload', () => {
        const saved = sessions.saveSession([
            { id: 'm-0', role: 'user', text: 'loadable', ts: 'x', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] },
            { id: 'm-1', role: 'assistant', text: 'answer', ts: 'y', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] },
        ]);
        const onLoadSession = jest.fn();
        initHistoryView({ onLoadSession, onDeleteSession: jest.fn(), onNewChat: jest.fn() });
        openHistory();

        document.querySelector('.history-item').click();

        expect(onLoadSession).toHaveBeenCalledTimes(1);
        expect(onLoadSession).toHaveBeenCalledWith(expect.objectContaining({ id: saved.id, messages: expect.any(Array) }));
    });

    test('clicking delete fires onDeleteSession and removes the row', () => {
        const a = sessions.saveSession([{ id: 'm-0', role: 'user', text: 'A', ts: 'x', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] }]);
        sessions.saveSession([{ id: 'm-0', role: 'user', text: 'B', ts: 'x', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] }]);

        const onDeleteSession = jest.fn();
        initHistoryView({ onLoadSession: jest.fn(), onDeleteSession, onNewChat: jest.fn() });
        openHistory();

        const deleteBtn = document.querySelector(`[data-session-id="${a.id}"] .history-item-delete`);
        deleteBtn.click();

        expect(onDeleteSession).toHaveBeenCalledTimes(1);
        expect(onDeleteSession).toHaveBeenCalledWith(a.id);
        // Row for the deleted session is gone; the other row remains.
        expect(document.querySelectorAll('.history-item')).toHaveLength(1);
    });

    test('clicking the overlay New chat button fires onNewChat', () => {
        sessions.saveSession([{ id: 'm-0', role: 'user', text: 'x', ts: 'x', status: '', error: null, worklog: null, model: null, citations: [], proposals: [] }]);
        const onNewChat = jest.fn();
        initHistoryView({ onLoadSession: jest.fn(), onDeleteSession: jest.fn(), onNewChat });
        openHistory();

        document.getElementById('historyNewChatBtn').click();

        expect(onNewChat).toHaveBeenCalledTimes(1);
    });
});

describe('closeHistory / isHistoryOpen', () => {
    test('closeHistory hides the overlay', () => {
        initHistoryView({ onLoadSession: jest.fn(), onDeleteSession: jest.fn(), onNewChat: jest.fn() });
        openHistory();
        expect(isHistoryOpen()).toBe(true);

        closeHistory();
        expect(isHistoryOpen()).toBe(false);
        expect(document.getElementById('historyOverlay').hasAttribute('hidden')).toBe(true);
    });

    test('clicking the overlay close button hides the overlay', () => {
        initHistoryView({ onLoadSession: jest.fn(), onDeleteSession: jest.fn(), onNewChat: jest.fn() });
        openHistory();

        document.getElementById('historyCloseBtn').click();
        expect(isHistoryOpen()).toBe(false);
    });
});
