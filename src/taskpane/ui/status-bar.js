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

/** Drawer DOM entries kept before the oldest are dropped. */
const MAX_LOG_ENTRIES = 200;

/**
 * Module state for the dev-server log endpoint probe. The POST is dev-only
 * convenience; in production /log 404s, so after the first failed probe the
 * client stops hammering the endpoint on every log line.
 */
let devLogEndpointAvailable = null; // null = not probed yet

/**
 * Best-effort POST of one log line to the dev-server /log endpoint. Probes
 * once: any failure (404 on production, dev server down, CORS) disables
 * further attempts for the session.
 *
 * @private
 */
function postToDevLog(body) {
    if (devLogEndpointAvailable === false) return;
    fetch('/log', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body,
    }).then((res) => {
        devLogEndpointAvailable = res.ok ? true : false;
    }).catch(() => {
        devLogEndpointAvailable = false;
    });
}

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
        // Long document runs log per chunk — cap the drawer so a whole
        // session's logs cannot grow without bound in the DOM.
        while (logsDiv.childElementCount > MAX_LOG_ENTRIES) {
            logsDiv.removeChild(logsDiv.firstElementChild);
        }
        logsDiv.scrollTop = logsDiv.scrollHeight;
    }

    console.log(`[${type.toUpperCase()}] ${message}`);

    // Send to server log (best effort)
    postToDevLog(JSON.stringify({ message, type, timestamp: new Date().toISOString() }));
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

    // Same probe guard as addLog: without it, retry logs fired a doomed POST
    // on every occurrence in production (the endpoint 404s there).
    postToDevLog(JSON.stringify({ message, type, timestamp: new Date().toISOString() }));
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
 * Updates the header connection indicator. The visible state is a colored
 * dot only (green = connected, yellow = connecting, red = error); the label
 * text is hidden and carried on the dot as a hover tooltip / aria-label so
 * the reason (e.g. "API key required") is not lost.
 *
 * @param {string} state - 'connecting' | 'connected' | 'error'
 * @param {string} text - Status label text
 */
export function setConnectionStatus(state, text) {
    const indicator = document.getElementById('statusIndicator');
    const statusText = document.getElementById('statusText');
    const container = document.getElementById('connectionStatus');
    if (indicator) {
        indicator.className = 'status-indicator';
        if (state === 'connected') indicator.classList.add('connected');
        if (state === 'error') indicator.classList.add('error');
    }
    if (statusText) {
        // Stored (span is visually hidden) so the tooltip text is reusable.
        statusText.textContent = text;
    }
    if (container) {
        container.title = text;
        container.setAttribute('aria-label', text);
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
    const closeBtn = document.getElementById('logDrawerCloseBtn');
    const clearBtn = document.getElementById('clearLogsBtn');

    // #logBtn is intentionally NOT bound here: the bootstrap wires it via
    // this module's exported toggleLogDrawer() so the click handler stays
    // in one place. Binding it here too made one click toggle the drawer
    // twice (open then instantly close).
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
