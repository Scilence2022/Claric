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
 * In addition to live DOM, the module keeps a parallel `currentSessionMessages`
 * array so the conversation can be serialized (history view, restart on
 * reload). Each call to `createAssistantMessage()` returns a handle whose
 * `finalizeForHistory()` snapshots the rendered state into the array. The
 * bootstrap wires a single `onTurnCommitted` callback that persists the
 * session after each turn settles.
 *
 * @module ui/chat-view
 */

import { buildTextDiffElement } from './diff-view.js';
import { renderTablePreview, sanitizeTablePreview } from './proposal-card.js';

let _messagesEl = null;
let _welcomeEl = null;

let _currentSessionId = null;
let _currentSessionTitle = null;
let _currentSessionCreatedAt = null;
let _currentSessionUpdatedAt = null;
let _currentSessionMessages = [];

// Fired when a proposal settles AFTER its message was finalized, so the
// bootstrap can re-persist the session (Apply/Reject arrive after the turn's
// finally block has already snapshotted the message).
let _proposalStateHandler = null;

/**
 * Registers the late proposal-state callback. The bootstrap wires this to
 * session persistence; tests can reset it with null.
 *
 * @param {function()|null} handler
 */
export function setProposalStateChangeHandler(handler) {
    _proposalStateHandler = typeof handler === 'function' ? handler : null;
}

/** Serializes one tracked proposal meta into its history shape. @private */
function _proposalRecordFromMeta(p) {
    return {
        title: p.title,
        state: p.state || 'pending',
        detail: p.detail,
        countsText: p.countsText,
        previewSrc: p.previewSrc || null,
        tablePreview: sanitizeTablePreview(p.tablePreview),
        items: Array.isArray(p.items) ? p.items.map((it) => ({
            id: it.id,
            label: it.label,
            before: it.before,
            after: it.after,
            searchText: it.searchText,
        })) : [],
    };
}

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
 * Removes all messages and restores the welcome state. Wipes the live
 * session tracking so the chat truly starts over.
 */
export function clearChat() {
    if (_messagesEl) {
        _messagesEl.querySelectorAll('.chat-message').forEach((el) => el.remove());
    }
    clearSessionMessages();
    showWelcome();
}

/**
 * Resets the live session tracking (messages array + ids). Leaves DOM alone.
 */
export function clearSessionMessages() {
    _currentSessionId = null;
    _currentSessionTitle = null;
    _currentSessionCreatedAt = null;
    _currentSessionUpdatedAt = null;
    _currentSessionMessages = [];
}

/**
 * Returns the live session object: id, title, timestamps, and the running
 * message array. The bootstrap persists a snapshot after each turn.
 *
 * @returns {{ id: string|null, title: string|null, createdAt: string|null, updatedAt: string|null, messages: Array<object> }}
 */
export function getCurrentSession() {
    return {
        id: _currentSessionId,
        title: _currentSessionTitle,
        createdAt: _currentSessionCreatedAt,
        updatedAt: _currentSessionUpdatedAt,
        messages: _currentSessionMessages,
    };
}

/**
 * Seeds the live session from a saved payload (history view "load" action).
 * The DOM is rebuilt via renderHistory() and the welcome state is hidden.
 *
 * @param {object} session
 * @param {string} [session.id]
 * @param {string} [session.title]
 * @param {string} [session.createdAt]
 * @param {string} [session.updatedAt]
 * @param {Array<object>} [session.messages]
 */
export function setCurrentSession(session) {
    if (!session) {
        clearSessionMessages();
        renderHistory([]);
        return;
    }
    _currentSessionId = session.id || _generateSessionId();
    _currentSessionTitle = session.title || null;
    _currentSessionCreatedAt = session.createdAt || new Date().toISOString();
    _currentSessionUpdatedAt = session.updatedAt || new Date().toISOString();
    _currentSessionMessages = Array.isArray(session.messages)
        ? session.messages.map(_normalizeMessage)
        : [];
    renderHistory(_currentSessionMessages);
}

/**
 * Builds DOM from a saved message array. Clears any existing messages,
 * hides the welcome state, and renders each message in order.
 *
 * @param {Array<object>} messages
 */
export function renderHistory(messages) {
    if (!_messagesEl) return;
    _messagesEl.querySelectorAll('.chat-message').forEach((el) => el.remove());
    if (!Array.isArray(messages) || messages.length === 0) {
        showWelcome();
        return;
    }
    hideWelcome();
    for (const m of messages) {
        if (!m) continue;
        if (m.role === 'user') {
            const bubble = document.createElement('div');
            bubble.className = 'chat-message chat-message-user';
            bubble.textContent = m.text || '';
            _messagesEl.appendChild(bubble);
        } else {
            _messagesEl.appendChild(_renderHistoricalAssistant(m));
        }
    }
    _scrollToBottom();
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
 * Renders a text node safely (no innerHTML).
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

    if (_currentSessionId === null) {
        _currentSessionId = _generateSessionId();
        _currentSessionCreatedAt = new Date().toISOString();
    }
    _currentSessionUpdatedAt = new Date().toISOString();
    _currentSessionMessages.push({
        id: _generateMessageId(),
        role: 'user',
        text: String(text || ''),
        status: '',
        error: null,
        worklog: null,
        model: null,
        citations: [],
        proposals: [],
        ts: new Date().toISOString(),
    });
}

/**
 * Adds a small system note (about info, hints) as an assistant-style message.
 */
export function addSystemNote(text) {
    hideWelcome();
    const msg = createAssistantMessage();
    msg.setText(text);
    return msg;
}

/**
 * Generates a stable id for the current session.
 * @private
 */
function _generateSessionId() {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `s-${crypto.randomUUID()}`;
    }
    return `s-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function _generateMessageId() {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `m-${crypto.randomUUID()}`;
    }
    return `m-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

/**
 * Normalizes one stored message into the live shape (drops unknown props).
 * @private
 */
function _normalizeMessage(m) {
    if (!m || typeof m !== 'object') return null;
    return {
        id: m.id || _generateMessageId(),
        role: m.role === 'assistant' ? 'assistant' : 'user',
        text: typeof m.text === 'string' ? m.text : '',
        status: typeof m.status === 'string' ? m.status : '',
        error: typeof m.error === 'string' ? m.error : null,
        worklog: m.worklog && typeof m.worklog === 'object'
            ? { count: Number(m.worklog.count) || 0, durationMs: Number(m.worklog.durationMs) || 0 }
            : null,
        model: m.model && typeof m.model === 'object'
            ? { sections: Number(m.model.sections) || 0 }
            : null,
        citations: Array.isArray(m.citations) ? m.citations.filter((c) => c && typeof c === 'object') : [],
        proposals: Array.isArray(m.proposals) ? m.proposals : [],
        ts: typeof m.ts === 'string' ? m.ts : new Date().toISOString(),
    };
}

/**
 * Wraps a proposal card so its terminal-state methods also update the meta
 * object that the chat-view tracks for history serialization. onStateChange
 * fires after every wrapped state change so a message that was already
 * finalized can re-sync its history record.
 * @private
 */
function _wrapProposalCard(card, meta, onStateChange) {
    const notify = typeof onStateChange === 'function' ? onStateChange : () => {};
    const origApplied = card.markApplied;
    const origRejected = card.markRejected;
    const origWarning = card.markWarning;
    const origError = card.markError;
    card.markApplied = function () {
        meta.state = 'applied';
        const result = origApplied.call(card);
        notify();
        return result;
    };
    card.markRejected = function () {
        meta.state = 'rejected';
        const result = origRejected.call(card);
        notify();
        return result;
    };
    card.markWarning = function (msg) {
        meta.state = 'warning';
        if (msg) meta.detail = String(msg);
        const result = origWarning.call(card, msg);
        notify();
        return result;
    };
    card.markError = function (msg) {
        meta.state = 'error';
        if (msg) meta.detail = String(msg);
        const result = origError.call(card, msg);
        notify();
        return result;
    };
}

/**
 * Renders an assistant message from saved history (no callbacks, no streaming).
 * @private
 */
function _renderHistoricalAssistant(m) {
    const el = document.createElement('div');
    el.className = 'chat-message chat-message-assistant';
    if (m.error) el.classList.add('chat-message-error');

    if (m.status) {
        const status = document.createElement('div');
        status.className = 'msg-status';
        status.textContent = m.status;
        el.appendChild(status);
    }

    if (m.worklog && Number(m.worklog.count) > 0) {
        const wrap = document.createElement('div');
        wrap.className = 'msg-worklog';
        wrap.style.display = '';
        const toggle = document.createElement('button');
        toggle.type = 'button';
        toggle.className = 'msg-worklog-toggle';
        toggle.setAttribute('aria-expanded', 'false');
        const secs = Math.max(1, Math.round((Number(m.worklog.durationMs) || 0) / 1000));
        const count = Number(m.worklog.count) || 0;
        toggle.textContent = `▸ Worked for ${secs}s · ${count} step${count === 1 ? '' : 's'}`;
        wrap.appendChild(toggle);
        el.appendChild(wrap);
    }

    if (m.model && Number(m.model.sections) > 0) {
        const wrap = document.createElement('div');
        wrap.className = 'msg-model';
        wrap.style.display = '';
        const toggle = document.createElement('button');
        toggle.type = 'button';
        toggle.className = 'msg-model-toggle';
        toggle.setAttribute('aria-expanded', 'false');
        const sections = Number(m.model.sections) || 0;
        toggle.textContent = `▸ Model activity · ${sections} section${sections === 1 ? '' : 's'}`;
        wrap.appendChild(toggle);
        el.appendChild(wrap);
    }

    const body = document.createElement('div');
    body.className = 'msg-body';
    body.textContent = m.error ? `Error: ${m.error}` : (m.text || '');
    el.appendChild(body);

    if (Array.isArray(m.proposals)) {
        for (const p of m.proposals) {
            const card = renderStaticProposalCard(p);
            if (card) el.appendChild(card);
        }
    }

    return el;
}

/**
 * Renders a read-only proposal card from saved history data. Shows terminal
 * state badges and per-change before/after diffs but no interactive controls.
 *
 * @param {object} p
 * @param {string} [p.title]
 * @param {string} [p.state] - 'applied' | 'rejected' | 'warning' | 'error'
 * @param {string} [p.detail] - extra status text for warning/error
 * @param {string} [p.countsText]
 * @param {string} [p.previewSrc]
 * @param {object} [p.tablePreview] - Sanitized read-only table preview data
 * @param {Array<object>} [p.items] - { label, before?, after?, searchText? }
 * @returns {HTMLElement|null}
 */
export function renderStaticProposalCard(p) {
    if (!p || typeof p !== 'object') return null;
    const el = document.createElement('div');
    el.className = 'proposal-card';
    if (p.state === 'applied') el.classList.add('proposal-applied');
    else if (p.state === 'rejected') el.classList.add('proposal-rejected');
    else if (p.state === 'warning') el.classList.add('proposal-warning');
    else if (p.state === 'error') el.classList.add('proposal-error');

    const head = document.createElement('div');
    head.className = 'proposal-card-head';
    const titleEl = document.createElement('span');
    titleEl.className = 'proposal-card-title';
    titleEl.textContent = p.title || 'Proposal';
    head.appendChild(titleEl);
    if (p.countsText) {
        const counts = document.createElement('span');
        counts.className = 'proposal-card-counts';
        counts.textContent = p.countsText;
        head.appendChild(counts);
    }
    el.appendChild(head);

    if (p.previewSrc) {
        const img = document.createElement('img');
        img.className = 'proposal-card-preview';
        img.alt = 'Proposal preview';
        img.src = p.previewSrc;
        el.appendChild(img);
    }

    const tablePreviewEl = renderTablePreview(p.tablePreview);
    if (tablePreviewEl) el.appendChild(tablePreviewEl);

    if (Array.isArray(p.items) && p.items.length) {
        const list = document.createElement('div');
        list.className = 'proposal-card-changes-static';
        for (const item of p.items) {
            if (!item) continue;
            const row = document.createElement('div');
            row.className = 'proposal-card-change';
            const label = document.createElement('div');
            label.className = 'proposal-card-change-label';
            label.textContent = item.label || '';
            row.appendChild(label);
            if (item.before !== undefined && item.after !== undefined) {
                row.appendChild(buildTextDiffElement(item.before, item.after));
            }
            list.appendChild(row);
        }
        el.appendChild(list);
    }

    if (p.state && p.state !== 'pending') {
        const status = document.createElement('div');
        status.className = 'proposal-card-status';
        let label;
        if (p.state === 'applied') label = 'Applied as tracked changes.';
        else if (p.state === 'rejected') label = 'Rejected — no changes were made.';
        else if (p.state === 'warning') label = p.detail || 'Nothing applied.';
        else if (p.state === 'error') label = `Apply failed: ${p.detail || ''}`;
        else label = p.state;
        status.textContent = label;
        status.style.display = '';
        el.appendChild(status);
    }

    return el;
}

/**
 * Creates an assistant message and returns a handle for updating it through
 * the turn lifecycle (status → streaming text/progress → final state).
 *
 * @returns {{
 *   el: HTMLElement,
 *   setStatus: function(string),
 *   setText: function(string),
 *   appendText: function(string),
 *   appendLogLine: function(string),
 *   collapseLog: function(),
 *   appendModelToken: function({id: string, index?: number}, string, string),
 *   collapseModelOutput: function(),
 *   showProgress: function(object),
 *   hideProgress: function(),
 *   attachProposal: function(card, meta?),
 *   addCitationPills: function(Array<{label: string, searchText: string}>, function(string)),
 *   markError: function(string),
 *   finalizeForHistory: function(),
 * }}
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

    // Collapsible model activity region: the model's streamed thinking
    // (dimmed) and output, split into per-chunk sections for document runs.
    const modelEl = document.createElement('div');
    modelEl.className = 'msg-model';
    modelEl.style.display = 'none';
    const modelToggle = document.createElement('button');
    modelToggle.type = 'button';
    modelToggle.className = 'msg-model-toggle';
    modelToggle.setAttribute('aria-expanded', 'true');
    const modelBody = document.createElement('div');
    modelBody.className = 'msg-model-body';
    modelEl.appendChild(modelToggle);
    modelEl.appendChild(modelBody);

    // Auto-scroll: the region follows the stream while the user stays near
    // the bottom; scrolling up (to read earlier output) disengages the
    // follow until they scroll back down.
    let modelStickToBottom = true;
    modelBody.addEventListener('scroll', () => {
        modelStickToBottom =
            modelBody.scrollHeight - modelBody.scrollTop - modelBody.clientHeight < 40;
    });

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
    el.appendChild(modelEl);
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
    const modelSections = new Map();
    let modelCollapsed = false;
    let lastStatus = '';
    let lastError = null;
    const trackedProposals = [];
    let finalized = false;
    let finalizedRecord = null;

    /**
     * Re-syncs the finalized history record when a proposal settles late
     * (Apply/Reject arrive after finalizeForHistory), then notifies the
     * bootstrap so the session is re-persisted.
     */
    function _syncFinalizedProposals() {
        if (!finalizedRecord) return;
        finalizedRecord.proposals = trackedProposals.map(_proposalRecordFromMeta);
        finalizedRecord.ts = new Date().toISOString();
        _currentSessionUpdatedAt = finalizedRecord.ts;
        if (_proposalStateHandler) {
            try {
                _proposalStateHandler();
            } catch (_err) {
                // Persistence errors must never break the live card.
            }
        }
    }

    if (_currentSessionId === null) {
        _currentSessionId = _generateSessionId();
        _currentSessionCreatedAt = new Date().toISOString();
    }

    worklogToggle.addEventListener('click', () => {
        const expanding = worklogLines.style.display === 'none';
        worklogLines.style.display = expanding ? '' : 'none';
        worklogToggle.setAttribute('aria-expanded', String(expanding));
        _renderWorklogToggle();
    });

    modelToggle.addEventListener('click', () => {
        const expanding = modelBody.style.display === 'none';
        modelBody.style.display = expanding ? '' : 'none';
        modelToggle.setAttribute('aria-expanded', String(expanding));
        if (expanding) {
            modelStickToBottom = true;
            modelBody.scrollTop = modelBody.scrollHeight;
        }
        _renderModelToggle();
    });

    function _renderModelToggle() {
        const expanded = modelBody.style.display !== 'none';
        const chevron = expanded ? '▾' : '▸';
        const n = modelSections.size;
        modelToggle.textContent = modelCollapsed
            ? `${chevron} Model activity · ${n} section${n === 1 ? '' : 's'}`
            : `${chevron} Model activity`;
    }

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
        setStatus(text) {
            lastStatus = text || '';
            statusEl.style.display = text ? '' : 'none';
            statusEl.textContent = text;
            _scrollToBottom();
        },
        setText(text) {
            streamed = text || '';
            _renderText(bodyEl, streamed);
            _scrollToBottom();
        },
        appendText(token) {
            streamed += token;
            _renderText(bodyEl, streamed);
            _scrollToBottom();
        },
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
        collapseLog() {
            if (logLineCount === 0) return;
            logCollapsed = true;
            worklogLines.style.display = 'none';
            worklogToggle.setAttribute('aria-expanded', 'false');
            _renderWorklogToggle();
        },
        appendModelToken(sectionRef, kind, token) {
            if (!token) return;
            const id = (sectionRef && sectionRef.id) || 'default';
            const index = sectionRef && typeof sectionRef.index === 'number' ? sectionRef.index : null;
            let section = modelSections.get(id);
            if (!section) {
                const sectionEl = document.createElement('div');
                sectionEl.className = 'msg-model-section';
                const labelEl = document.createElement('div');
                labelEl.className = 'msg-model-section-label';
                labelEl.textContent = index !== null ? `Section ${index + 1}` : '';
                labelEl.style.display = 'none';
                sectionEl.appendChild(labelEl);
                modelBody.appendChild(sectionEl);
                section = { el: sectionEl, labelEl, reasoningEl: null, contentEl: null };
                modelSections.set(id, section);
                if (modelSections.size > 1) {
                    for (const s of modelSections.values()) {
                        if (s.labelEl.textContent) s.labelEl.style.display = '';
                    }
                }
                modelEl.style.display = '';
                _renderModelToggle();
            }
            let target;
            if (kind === 'reasoning') {
                if (!section.reasoningEl) {
                    const r = document.createElement('div');
                    r.className = 'msg-model-reasoning';
                    section.el.insertBefore(r, section.contentEl || null);
                    section.reasoningEl = r;
                }
                target = section.reasoningEl;
            } else {
                if (!section.contentEl) {
                    const c = document.createElement('div');
                    c.className = 'msg-model-content';
                    section.el.appendChild(c);
                    section.contentEl = c;
                }
                target = section.contentEl;
            }
            target.textContent += token;
            if (!modelCollapsed) {
                if (modelStickToBottom) {
                    modelBody.scrollTop = modelBody.scrollHeight;
                }
                _scrollToBottom();
            }
        },
        collapseModelOutput() {
            if (modelSections.size === 0) return;
            modelCollapsed = true;
            modelBody.style.display = 'none';
            modelToggle.setAttribute('aria-expanded', 'false');
            _renderModelToggle();
        },
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
        hideProgress() {
            progressEl.style.display = 'none';
        },
        attachProposal(card, meta) {
            if (meta) {
                _wrapProposalCard(card, meta, _syncFinalizedProposals);
                trackedProposals.push(meta);
            }
            extrasEl.appendChild(card.el);
            _scrollToBottom();
        },
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
        markError(message) {
            lastError = String(message || '');
            streamed = `Error: ${lastError}`;
            el.classList.add('chat-message-error');
            statusEl.style.display = 'none';
            bodyEl.textContent = streamed;
            _scrollToBottom();
        },
        /**
         * Snapshots the current message state into the live session array.
         * Idempotent — repeated calls are no-ops (so a turn that finalizes
         * via both collapseLog and an explicit caller only commits once).
         */
        finalizeForHistory() {
            if (finalized) return;
            finalized = true;
            _currentSessionUpdatedAt = new Date().toISOString();
            const now = new Date().toISOString();
            finalizedRecord = {
                id: _generateMessageId(),
                role: 'assistant',
                text: streamed || '',
                status: lastStatus || '',
                error: lastError,
                worklog: logLineCount > 0
                    ? { count: logLineCount, durationMs: Math.max(0, Date.now() - logStartTime) }
                    : null,
                model: modelSections.size > 0 ? { sections: modelSections.size } : null,
                citations: [],
                proposals: trackedProposals.map(_proposalRecordFromMeta),
                ts: now,
            };
            _currentSessionMessages.push(finalizedRecord);
        },
    };
}
