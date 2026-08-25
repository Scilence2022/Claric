/**
 * Status Bar Module
 *
 * Owns the floating activity-log drawer, the comment-pending status bar, and
 * the header connection indicator. The activity log keeps the original
 * behavior (timestamped entries + best-effort dev-server POST /log), but
 * defaults to hidden and is opened/closed by the header #logBtn.
 *
 * @module ui/status-bar
 */

/**
 * Appends an entry to the activity log drawer.
 *
 * @param {string} message - Log message text
 * @param {string} [type='info'] - One of 'info' | 'success' | 'warning' | 'error'
 */
export function addLog(message, type = 'info') {
    const logsDiv = typeof document !== 'undefined' ? document.getElementById('logs') : null;
    if (logsDiv) {
        const entry = document.createElement('div');
        const timestamp = new Date().toLocaleTimeString();

        entry.className = `log-${type}`;
        entry.textContent = `[${timestamp}] ${message}`;

        logsDiv.appendChild(entry);
        logsDiv.scrollTop = logsDiv.scrollHeight;
    }

    console.log(`[${type.toUpperCase()}] ${message}`);

    // Send to server log (best effort)
    fetch('/log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message, type, timestamp: new Date().toISOString() })
    }).catch(() => { });
}

/**
 * Extended version of addLog that appends a clickable "Retry" link to the
 * log entry. Used for failed comment/chunk requests the user can retry.
 *
 * @param {string} message - The log message text
 * @param {string} type - Log type: 'info', 'success', 'warning', 'error'
 * @param {Function} retryCallback - Function to call when Retry is clicked
 */
export function addLogWithRetry(message, type, retryCallback) {
    const logsDiv = typeof document !== 'undefined' ? document.getElementById('logs') : null;
    if (!logsDiv) {
        addLog(message, type);
        return;
    }

    const entry = document.createElement('div');
    const timestamp = new Date().toLocaleTimeString();
    entry.className = `log-${type}`;

    const msgSpan = document.createElement('span');
    msgSpan.textContent = `[${timestamp}] ${message} `;
    entry.appendChild(msgSpan);

    if (retryCallback) {
        const retryLink = document.createElement('a');
        retryLink.textContent = 'Retry';
        retryLink.href = '#';
        retryLink.className = 'retry-link';
        retryLink.onclick = (e) => {
            e.preventDefault();
            retryCallback();
            entry.remove();  // Remove the error log entry on retry
        };
        entry.appendChild(retryLink);
    }

    logsDiv.appendChild(entry);
    logsDiv.scrollTop = logsDiv.scrollHeight;

    console.log(`[${type.toUpperCase()}] ${message}`);

    fetch('/log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message, type, timestamp: new Date().toISOString() })
    }).catch(() => { });
}

/**
 * Clears all entries from the activity log.
 */
export function clearLogs() {
    const logsDiv = document.getElementById('logs');
    if (logsDiv) logsDiv.innerHTML = '';
}

/**
 * Updates the comment-pending status bar above the input area.
 *
 * @param {number} count - Number of comment requests currently pending
 */
export function updateCommentStatusBar(count) {
    const bar = document.getElementById('commentStatusBar');
    if (!bar) return;

    if (count === 0) {
        bar.style.display = 'none';
    } else {
        bar.style.display = 'flex';
        const text = document.getElementById('commentStatusText');
        if (text) {
            text.textContent = `${count} comment${count !== 1 ? 's' : ''} pending...`;
        }
    }
}

/**
 * Updates the header connection indicator.
 *
 * @param {string} state - 'connecting' | 'connected' | 'error'
 * @param {string} text - Status label text
 */
export function setConnectionStatus(state, text) {
    const indicator = document.getElementById('statusIndicator');
    const statusText = document.getElementById('statusText');
    if (indicator) {
        indicator.className = 'status-indicator';
        if (state === 'connected') indicator.classList.add('connected');
        if (state === 'error') indicator.classList.add('error');
    }
    if (statusText) {
        statusText.textContent = text;
    }
}

/**
 * Wires the floating log drawer to the header button, the drawer's close
 * button, and the Clear control. Called once at startup. Drawer is hidden
 * by default; the caller wires #logBtn via this module's exported
 * toggleLogDrawer() so the click handler stays in one place.
 */
export function initStatusBar() {
    const drawer = document.getElementById('logDrawer');
    const logBtn = document.getElementById('logBtn');
    const closeBtn = document.getElementById('logDrawerCloseBtn');
    const clearBtn = document.getElementById('clearLogsBtn');

    if (logBtn && drawer) {
        logBtn.addEventListener('click', () => toggleLogDrawer());
    }
    if (closeBtn && drawer) {
        closeBtn.addEventListener('click', () => closeLogDrawer());
    }
    if (clearBtn) {
        clearBtn.addEventListener('click', clearLogs);
    }
}

/**
 * Toggles the floating log drawer open/closed. Returns true when the drawer
 * is open after the call.
 *
 * @returns {boolean}
 */
export function toggleLogDrawer() {
    const drawer = document.getElementById('logDrawer');
    if (!drawer) return false;
    const willOpen = drawer.hasAttribute('hidden');
    if (willOpen) {
        drawer.removeAttribute('hidden');
    } else {
        drawer.setAttribute('hidden', '');
    }
    const btn = document.getElementById('logBtn');
    if (btn) btn.setAttribute('aria-expanded', String(willOpen));
    return willOpen;
}

/**
 * Forces the log drawer closed. Safe to call when already closed.
 */
export function closeLogDrawer() {
    const drawer = document.getElementById('logDrawer');
    if (!drawer) return;
    drawer.setAttribute('hidden', '');
    const btn = document.getElementById('logBtn');
    if (btn) btn.setAttribute('aria-expanded', 'false');
}

/**
 * True when the floating log drawer is currently visible.
 * @returns {boolean}
 */
export function isLogDrawerOpen() {
    const drawer = document.getElementById('logDrawer');
    return !!drawer && !drawer.hasAttribute('hidden');
}
