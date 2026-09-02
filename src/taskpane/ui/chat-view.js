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

import { sanitizeSvg } from '../../lib/illustration.js';
import { buildTextDiffElement } from './diff-view.js';
import { appState } from '../app-state.js';
import { renderTablePreview, sanitizeTablePreview } from './proposal-card.js';
import { newId, normalizeAttachments, normalizeCitations, normalizeMessage } from '../message-shape.js';

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

// Reveal callback for citation pills rebuilt from history. A stored message
// carries no closures, so the bootstrap supplies the "select this text in the
// document" action once and every restored pill uses it.
let _citationSelectHandler = null;

/**
 * Registers the reveal action for citation pills rendered from history.
 * Without it, restored pills are omitted rather than rendered inert.
 *
 * @param {function(string)|null} handler - Called with a pill's searchText
 */
export function setCitationSelectHandler(handler) {
    _citationSelectHandler = typeof handler === 'function' ? handler : null;
}

/**
 * Builds the citation pill row. Shared by the live turn (addCitationPills)
 * and the history render, so both look and behave the same.
 *
 * @param {Array<{label: string, searchText: string}>} citations
 * @param {function(string)} onSelect
 * @returns {HTMLElement}
 * @private
 */
function _buildCitationRow(citations, onSelect) {
    const row = document.createElement('div');
    row.className = 'citation-row';
    for (const c of citations) {
        const pill = document.createElement('button');
        pill.className = 'citation-pill';
        pill.type = 'button';
        pill.textContent = `§ ${c.label}`;
        pill.title = c.searchText;
        pill.addEventListener('click', () => onSelect(c.searchText));
        row.appendChild(pill);
    }
    return row;
}

/**
 * Renders a static one-line summary block (worklog / model activity) for a
 * message restored from history, where only the counts were persisted.
 *
 * @param {string} className - 'msg-worklog' | 'msg-model'
 * @param {string} text
 * @returns {HTMLElement}
 * @private
 */
function _renderStaticSummary(className, text) {
    const wrap = document.createElement('div');
    wrap.className = className;
    const line = document.createElement('div');
    line.className = `${className}-summary`;
    line.textContent = text;
    wrap.appendChild(line);
    return wrap;
}

/** Serializes one tracked proposal meta into its history shape. @private */
function _proposalRecordFromMeta(p) {
    return {
        title: p.title,
        state: p.state || 'pending',
        detail: p.detail,
        countsText: p.countsText,
        previewSrc: p.previewSrc || null,
        previewSvg: p.previewSvg || null,
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
            const meta = _normalizeAttachments(m.attachments);
            if (meta.length > 0) {
                bubble.appendChild(_renderAttachmentMarkers(meta));
            }
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
        _queueScroll(_messagesEl);
    }
}

// Streaming callbacks fire per token; every scrollTop write paired with a
// scrollHeight read forces a layout, so a long stream reflowed the pane
// hundreds of times per second. Coalesce to one scroll write per animation
// frame. Falls back to immediate writes where requestAnimationFrame is
// unavailable (jsdom tests rely on the synchronous write).
const _pendingScrolls = new Map();
function _queueScroll(el) {
    if (typeof requestAnimationFrame !== 'function') {
        el.scrollTop = el.scrollHeight;
        return;
    }
    if (_pendingScrolls.has(el)) return;
    _pendingScrolls.set(el, requestAnimationFrame(() => {
        _pendingScrolls.delete(el);
        el.scrollTop = el.scrollHeight;
    }));
}

/**
 * Appends text incrementally: a text-node append is O(token), while
 * `el.textContent = fullString` rewrites the whole accumulated text per
 * token (O(n²) across a long stream). Falls back to a full render when the
 * element holds non-text content.
 * @private
 */
function _appendStreamText(el, text) {
    if (el.childNodes.length === 0) {
        el.appendChild(document.createTextNode(text));
        return;
    }
    if (el.childNodes.length === 1 && el.firstChild.nodeType === Node.TEXT_NODE) {
        el.firstChild.appendData(text);
        return;
    }
    el.textContent += text;
}

/**
 * Renders a text node safely (no innerHTML).
 * @private
 */
function _renderText(el, text) {
    el.textContent = text;
}

/**
 * Adds a user message bubble. attachments (optional) renders a row of file
 * markers under the text and is snapshotted into the session metadata
 * (name/kind/size only — never file bytes or data URLs).
 *
 * @param {string} text
 * @param {Array<{name: string, kind: string, size: number}>} [attachments]
 */
export function addUserMessage(text, attachments) {
    const meta = _normalizeAttachments(attachments);
    const el = document.createElement('div');
    el.className = 'chat-message chat-message-user';
    _renderText(el, text);
    if (meta.length > 0) {
        el.appendChild(_renderAttachmentMarkers(meta));
    }
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
        attachments: meta,
        ts: new Date().toISOString(),
    });
}

/** Keeps only display metadata from an attachment list. @private */
function _normalizeAttachments(attachments) {
    return normalizeAttachments(attachments);
}

/** Renders the file-marker row under a user bubble. @private */
function _renderAttachmentMarkers(attachments) {
    const row = document.createElement('div');
    row.className = 'user-attachments';
    for (const att of attachments) {
        const chip = document.createElement('span');
        chip.className = 'user-attachment';
        chip.textContent = `📎 ${att.name}`;
        chip.title = `${att.name} (${att.kind})`;
        row.appendChild(chip);
    }
    return row;
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
    return newId('s');
}

function _generateMessageId() {
    return newId('m');
}

/**
 * Normalizes one stored message into the live shape (drops unknown props).
 * Null entries are filtered by the callers.
 * @private
 */
function _normalizeMessage(m) {
    if (!m || typeof m !== 'object') return null;
    return normalizeMessage(m);
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

    // Only the counts are persisted (the log lines and streamed model output
    // are not), so history shows a static summary. Rendering a collapsed
    // toggle here would promise an expansion that has no content behind it.
    if (m.worklog && Number(m.worklog.count) > 0) {
        const secs = Math.max(1, Math.round((Number(m.worklog.durationMs) || 0) / 1000));
        const count = Number(m.worklog.count) || 0;
        el.appendChild(_renderStaticSummary(
            'msg-worklog', `Worked for ${secs}s · ${count} step${count === 1 ? '' : 's'}`));
    }

    if (m.model && Number(m.model.sections) > 0) {
        const sections = Number(m.model.sections) || 0;
        el.appendChild(_renderStaticSummary(
            'msg-model', `Model activity · ${sections} section${sections === 1 ? '' : 's'}`));
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

    // Citation pills survive a reload: the searchText is persisted, and the
    // click handler is re-supplied by the bootstrap (setCitationSelectHandler)
    // since a stored message carries no closures.
    const citations = normalizeCitations(m.citations);
    if (citations.length > 0 && _citationSelectHandler) {
        el.appendChild(_buildCitationRow(citations, _citationSelectHandler));
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
 * @param {string} [p.previewSvg] - Sanitized SVG markup, rendered inline
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

    if (p.previewSvg) {
        // Inline sanitized SVG — same host-agnostic reasoning as the live
        // proposal card: SVG data URLs fail to decode on some hosts.
        const holder = document.createElement('div');
        holder.className = 'proposal-card-preview proposal-card-preview-svg';
        holder.innerHTML = sanitizeSvg(p.previewSvg);
        el.appendChild(holder);
    } else if (p.previewSrc) {
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
    /** Citation pills added this turn — snapshotted by finalizeForHistory. */
    const trackedCitations = [];
    /** Cards staged with auto-apply on — drained by finalizeForHistory. */
    const pendingAutoApplyCards = [];
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
            _appendStreamText(bodyEl, token);
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
            // A chunked document run logs per chunk — cap the DOM rows per
            // message (the counter still reflects the true step count).
            while (worklogLines.childElementCount > 100) {
                worklogLines.removeChild(worklogLines.firstElementChild);
            }
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
            // Incremental append: `target.textContent += token` re-serializes
            // the whole accumulated text node per token (O(n²) over a long
            // stream); appending a text node touches only the new token.
            target.appendChild(document.createTextNode(token));
            if (!modelCollapsed) {
                if (modelStickToBottom) {
                    _queueScroll(modelBody);
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
            // Auto-apply mode: stage the card now, fire its apply once the
            // turn settles (finalizeForHistory) so the write never races
            // the turn's own busy-flag teardown.
            if (
                appState.config.autoApplyChanges === true &&
                meta && meta.state === 'pending' &&
                typeof card.applyAll === 'function'
            ) {
                pendingAutoApplyCards.push(card);
            }
        },
        addCitationPills(citations, onSelect) {
            const list = normalizeCitations(citations);
            if (list.length === 0) return;
            // Recorded for finalizeForHistory: the pills are rebuilt on
            // reload from the persisted label/searchText.
            trackedCitations.push(...list);
            extrasEl.appendChild(_buildCitationRow(list, onSelect));
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
            if (pendingAutoApplyCards.length > 0) {
                const cardsToApply = pendingAutoApplyCards.splice(0);
                // Sequential drain: a card's apply owns the cross-card
                // mutex and the document busy flags while in flight, so
                // concurrent applyAll calls would refuse each other.
                setTimeout(async () => {
                    for (const cardToApply of cardsToApply) {
                        try {
                            await cardToApply.applyAll();
                        } catch (_err) {
                            // onApply reports its own failure on the card.
                        }
                    }
                }, 0);
            }
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
                citations: trackedCitations.slice(),
                proposals: trackedProposals.map(_proposalRecordFromMeta),
                ts: now,
            };
            _currentSessionMessages.push(finalizedRecord);
        },
    };
}
