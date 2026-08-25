/**
 * Chat View Module
 *
 * Renders the message list: user bubbles, assistant messages with status
 * lines, streaming text, progress bars, proposal cards, and citation pills.
 * Also owns the welcome empty state container.
 *
 * All text is rendered via textContent (no innerHTML for model output), so
 * LLM responses cannot inject markup into the taskpane.
 *
 * @module ui/chat-view
 */

let _messagesEl = null;
let _welcomeEl = null;

/**
 * Captures the chat containers. Called once at startup.
 */
export function initChatView() {
    _messagesEl = document.getElementById('chatMessages');
    _welcomeEl = document.getElementById('welcome');
}

/**
 * Hides the welcome empty state (first turn of a conversation).
 */
export function hideWelcome() {
    if (_welcomeEl) _welcomeEl.style.display = 'none';
}

/**
 * Shows the welcome empty state (after clearing the chat).
 */
export function showWelcome() {
    if (_welcomeEl) _welcomeEl.style.display = '';
}

/**
 * Removes all messages and restores the welcome state.
 */
export function clearChat() {
    if (_messagesEl) {
        _messagesEl.querySelectorAll('.chat-message').forEach((el) => el.remove());
    }
    showWelcome();
}

/**
 * Scrolls the message list to the bottom.
 * @private
 */
function _scrollToBottom() {
    if (_messagesEl) {
        _messagesEl.scrollTop = _messagesEl.scrollHeight;
    }
}

/**
 * Renders text into an element as plain text with line breaks preserved.
 * @private
 */
function _renderText(el, text) {
    el.textContent = text;
}

/**
 * Adds a user message bubble.
 *
 * @param {string} text
 */
export function addUserMessage(text) {
    const el = document.createElement('div');
    el.className = 'chat-message chat-message-user';
    _renderText(el, text);
    _messagesEl.appendChild(el);
    _scrollToBottom();
}

/**
 * Adds a small system note (about info, hints) as an assistant-style message.
 *
 * @param {string} text
 */
export function addSystemNote(text) {
    hideWelcome();
    const msg = createAssistantMessage();
    msg.setText(text);
    return msg;
}

/**
 * Creates an assistant message and returns a handle for updating it through
 * the turn lifecycle (status → streaming text/progress → final state).
 *
 * The message includes a collapsible "work log" region (Claude Code style):
 * pipeline log lines stream into it while the turn runs (expanded), and it
 * auto-collapses to a one-line summary ("Worked for Ns · M steps") when the
 * turn finishes via collapseLog().
 *
 * @returns {{
 *   el: HTMLElement,
 *   setStatus: function(string),
 *   setText: function(string),
 *   appendText: function(string),
 *   appendLogLine: function(string),
 *   collapseLog: function(),
 *   showProgress: function(object),
 *   hideProgress: function(),
 *   attachProposal: function(HTMLElement),
 *   addCitationPills: function(Array<{label: string, searchText: string}>, function(string)),
 *   markError: function(string),
 * }} Handle for the new message
 */
export function createAssistantMessage() {
    const el = document.createElement('div');
    el.className = 'chat-message chat-message-assistant';

    // Collapsible work log (expanded while the turn runs)
    const worklogEl = document.createElement('div');
    worklogEl.className = 'msg-worklog';
    worklogEl.style.display = 'none';
    const worklogToggle = document.createElement('button');
    worklogToggle.type = 'button';
    worklogToggle.className = 'msg-worklog-toggle';
    worklogToggle.setAttribute('aria-expanded', 'true');
    const worklogLines = document.createElement('div');
    worklogLines.className = 'msg-worklog-lines';
    worklogEl.appendChild(worklogToggle);
    worklogEl.appendChild(worklogLines);

    const statusEl = document.createElement('div');
    statusEl.className = 'msg-status';
    statusEl.style.display = 'none';

    const bodyEl = document.createElement('div');
    bodyEl.className = 'msg-body';

    const progressEl = document.createElement('div');
    progressEl.className = 'msg-progress';
    progressEl.style.display = 'none';
    const progressTrack = document.createElement('div');
    progressTrack.className = 'progress-track';
    const progressFill = document.createElement('div');
    progressFill.className = 'progress-fill';
    progressTrack.appendChild(progressFill);
    const progressText = document.createElement('span');
    progressText.className = 'progress-text';
    progressEl.appendChild(progressTrack);
    progressEl.appendChild(progressText);

    const extrasEl = document.createElement('div');
    extrasEl.className = 'msg-extras';

    el.appendChild(worklogEl);
    el.appendChild(statusEl);
    el.appendChild(bodyEl);
    el.appendChild(progressEl);
    el.appendChild(extrasEl);
    _messagesEl.appendChild(el);
    _scrollToBottom();

    let streamed = '';
    let logLineCount = 0;
    let logStartTime = 0;
    let logCollapsed = false;

    worklogToggle.addEventListener('click', () => {
        const expanding = worklogLines.style.display === 'none';
        worklogLines.style.display = expanding ? '' : 'none';
        worklogToggle.setAttribute('aria-expanded', String(expanding));
        _renderWorklogToggle();
    });

    function _renderWorklogToggle() {
        const expanded = worklogLines.style.display !== 'none';
        const chevron = expanded ? '▾' : '▸';
        if (logCollapsed) {
            const secs = Math.max(1, Math.round((Date.now() - logStartTime) / 1000));
            worklogToggle.textContent = `${chevron} Worked for ${secs}s · ${logLineCount} step${logLineCount === 1 ? '' : 's'}`;
        } else {
            worklogToggle.textContent = `${chevron} Working…`;
        }
    }

    return {
        el,
        /** Sets the small status line above the message body ('' hides it). */
        setStatus(text) {
            statusEl.style.display = text ? '' : 'none';
            statusEl.textContent = text;
            _scrollToBottom();
        },
        /** Replaces the message body text. */
        setText(text) {
            streamed = text || '';
            _renderText(bodyEl, streamed);
            _scrollToBottom();
        },
        /** Appends a streamed token to the message body. */
        appendText(token) {
            streamed += token;
            _renderText(bodyEl, streamed);
            _scrollToBottom();
        },
        /** Appends one pipeline log line to the expanded work log. */
        appendLogLine(text) {
            if (logLineCount === 0) {
                logStartTime = Date.now();
                worklogEl.style.display = '';
                worklogLines.style.display = '';
                _renderWorklogToggle();
            }
            logLineCount++;
            const line = document.createElement('div');
            line.className = 'msg-worklog-line';
            line.textContent = text;
            worklogLines.appendChild(line);
            if (!logCollapsed) _scrollToBottom();
        },
        /** Collapses the work log to a one-line duration summary. */
        collapseLog() {
            if (logLineCount === 0) return;
            logCollapsed = true;
            worklogLines.style.display = 'none';
            worklogToggle.setAttribute('aria-expanded', 'false');
            _renderWorklogToggle();
        },
        /** Shows/updates the chunk progress bar. */
        showProgress(p) {
            progressEl.style.display = '';
            progressFill.style.width = `${p.percentComplete}%`;
            let label = `Processing: ${p.completed + p.failed + (p.cancelled || 0)}/${p.total} chunks`;
            if (p.estimatedSecondsRemaining > 0) {
                label += ` (~${p.estimatedSecondsRemaining}s remaining)`;
            }
            progressText.textContent = label;
            _scrollToBottom();
        },
        /** Hides the chunk progress bar. */
        hideProgress() {
            progressEl.style.display = 'none';
        },
        /** Attaches a proposal card element under the message. */
        attachProposal(cardEl) {
            extrasEl.appendChild(cardEl);
            _scrollToBottom();
        },
        /**
         * Adds citation pill buttons (one per processed chunk); clicking a
         * pill calls onSelect with the chunk's search text.
         */
        addCitationPills(citations, onSelect) {
            if (!citations || citations.length === 0) return;
            const row = document.createElement('div');
            row.className = 'citation-row';
            citations.forEach((c) => {
                const pill = document.createElement('button');
                pill.className = 'citation-pill';
                pill.type = 'button';
                pill.textContent = `§ ${c.label}`;
                pill.title = c.searchText;
                pill.addEventListener('click', () => onSelect(c.searchText));
                row.appendChild(pill);
            });
            extrasEl.appendChild(row);
            _scrollToBottom();
        },
        /** Marks the message as failed with an error line. */
        markError(message) {
            el.classList.add('chat-message-error');
            statusEl.style.display = 'none';
            bodyEl.textContent = `Error: ${message}`;
            _scrollToBottom();
        },
    };
}
