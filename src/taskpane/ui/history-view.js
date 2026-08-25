/**
 * History View Module
 *
 * Renders the chat-history slide-over: a list of saved sessions sorted
 * newest-first, with a click-to-load row action, a per-row delete button,
 * and a New chat button. Owns no persistence — `sessions.js` does that.
 *
 * @module ui/history-view
 */

import { listSessions, loadSession } from '../sessions.js';

let _closeBtnEl = null;
let _newChatBtnEl = null;

let _onLoadSession = null;
let _onDeleteSession = null;
let _onNewChat = null;

/**
 * Resolves the live overlay + body elements. Always re-queries the DOM so
 * the module survives jsdom resetting the document between tests, and so a
 * caller that mounts the overlay later still works.
 * @private
 */
function _resolveEls() {
    return {
        overlay: document.getElementById('historyOverlay'),
        body: document.getElementById('historyBody'),
    };
}

/**
 * One-time wiring. Stores the callbacks and binds the overlay's static
 * controls (close button, New chat button). The list itself is rebuilt
 * every time openHistory() runs.
 *
 * @param {object} deps
 * @param {function(object)} deps.onLoadSession - Called with the full session
 *   payload when the user clicks a row. Bootstrap reloads chat-view with it.
 * @param {function(string)} deps.onDeleteSession - Called with the session id
 *   when the user clicks the row's Delete button. Bootstrap is responsible
 *   for the actual delete + UI refresh.
 * @param {function()} [deps.onNewChat] - Called when the user clicks the
 *   overlay's New chat button. Bootstrap clears the live chat.
 */
export function initHistoryView({ onLoadSession, onDeleteSession, onNewChat } = {}) {
    _onLoadSession = typeof onLoadSession === 'function' ? onLoadSession : null;
    _onDeleteSession = typeof onDeleteSession === 'function' ? onDeleteSession : null;
    _onNewChat = typeof onNewChat === 'function' ? onNewChat : null;

    _closeBtnEl = document.getElementById('historyCloseBtn');
    _newChatBtnEl = document.getElementById('historyNewChatBtn');

    if (_closeBtnEl) {
        _closeBtnEl.addEventListener('click', () => closeHistory());
    }
    if (_newChatBtnEl) {
        _newChatBtnEl.addEventListener('click', () => {
            closeHistory();
            if (_onNewChat) _onNewChat();
        });
    }
}

/**
 * Opens the overlay and renders the current session list. Returns the
 * rendered session metadata for callers that want to inspect it (tests).
 *
 * @returns {Array<object>}
 */
export function openHistory() {
    const { overlay, body } = _resolveEls();
    if (!overlay || !body) return [];

    const sessions = listSessions();
    body.innerHTML = '';

    if (sessions.length === 0) {
        const empty = document.createElement('div');
        empty.className = 'history-empty';
        empty.textContent = 'No saved conversations yet.';
        body.appendChild(empty);
    } else {
        const list = document.createElement('div');
        list.className = 'history-list';
        for (const meta of sessions) {
            list.appendChild(_buildRow(meta));
        }
        body.appendChild(list);
    }

    overlay.removeAttribute('hidden');
    return sessions;
}

/**
 * Closes the overlay. Safe to call when already closed.
 */
export function closeHistory() {
    const { overlay } = _resolveEls();
    if (!overlay) return;
    overlay.setAttribute('hidden', '');
}

/**
 * True when the history overlay is currently visible.
 * @returns {boolean}
 */
export function isHistoryOpen() {
    const { overlay } = _resolveEls();
    return !!overlay && !overlay.hasAttribute('hidden');
}

/**
 * Rebuilds one session row with title, preview, relative-time meta, and
 * Delete button. Wired with the callbacks captured at init time.
 * @private
 */
function _buildRow(meta) {
    const row = document.createElement('div');
    row.className = 'history-item';
    row.dataset.sessionId = meta.id;
    row.setAttribute('role', 'button');
    row.setAttribute('tabindex', '0');
    row.title = meta.title || 'Untitled chat';

    const main = document.createElement('div');
    main.className = 'history-item-main';

    const title = document.createElement('div');
    title.className = 'history-item-title';
    title.textContent = meta.title || 'Untitled chat';
    main.appendChild(title);

    if (meta.preview) {
        const preview = document.createElement('div');
        preview.className = 'history-item-preview';
        preview.textContent = meta.preview;
        main.appendChild(preview);
    }

    const metaEl = document.createElement('div');
    metaEl.className = 'history-item-meta';
    const ts = document.createElement('span');
    ts.textContent = _formatRelativeTime(meta.updatedAt);
    metaEl.appendChild(ts);
    const sep = document.createElement('span');
    sep.textContent = '·';
    metaEl.appendChild(sep);
    const count = document.createElement('span');
    const msgCount = meta.messageCount || 0;
    count.textContent = `${msgCount} message${msgCount === 1 ? '' : 's'}`;
    metaEl.appendChild(count);
    main.appendChild(metaEl);

    row.appendChild(main);

    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'history-item-delete';
    del.textContent = 'Delete';
    del.title = 'Delete this conversation';
    del.setAttribute('aria-label', `Delete ${meta.title || 'conversation'}`);
    del.addEventListener('click', (e) => {
        e.stopPropagation();
        if (_onDeleteSession) _onDeleteSession(meta.id);
        // Optimistic remove: the caller may also re-render or delete via
        // sessions.deleteSession — either way the row should disappear now.
        const parent = row.parentNode;
        if (parent) parent.removeChild(row);
        // Empty body → show empty state inline.
        const { body } = _resolveEls();
        if (body && body.querySelectorAll('.history-item').length === 0) {
            const empty = document.createElement('div');
            empty.className = 'history-empty';
            empty.textContent = 'No saved conversations yet.';
            body.innerHTML = '';
            body.appendChild(empty);
        }
    });
    row.appendChild(del);

    row.addEventListener('click', () => {
        const full = loadSession(meta.id);
        if (full && _onLoadSession) {
            _onLoadSession(full);
            closeHistory();
        }
    });
    row.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            row.click();
        }
    });

    return row;
}

/**
 * Formats a timestamp like "5 minutes ago" / "yesterday". Returns '' for
 * missing or unparseable timestamps.
 * @private
 */
function _formatRelativeTime(iso) {
    if (!iso) return '';
    const then = new Date(iso).getTime();
    if (!Number.isFinite(then)) return '';
    const diffMs = Date.now() - then;
    const diffSec = Math.round(diffMs / 1000);
    if (Math.abs(diffSec) < 60) return 'just now';
    const formatter = new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' });
    if (Math.abs(diffSec) < 3600) {
        return formatter.format(-Math.round(diffSec / 60), 'minute');
    }
    if (Math.abs(diffSec) < 86400) {
        return formatter.format(-Math.round(diffSec / 3600), 'hour');
    }
    if (Math.abs(diffSec) < 86400 * 7) {
        return formatter.format(-Math.round(diffSec / 86400), 'day');
    }
    return new Date(iso).toLocaleDateString();
}

/**
 * Test seam.
 */
export const __testing = {
    _formatRelativeTime,
};
