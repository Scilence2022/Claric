/**
 * Sessions storage layer.
 *
 * Persists chat sessions under the selected workspace/document namespace.
 * With no scope selected, the legacy wordAI.sessions.index and wordAI.session.*
 * keys remain accessible; scoped history never migrates or deletes them.
 *
 * Goals:
 *   - Append-only-friendly writes (update bumps updatedAt + keeps createdAt).
 *   - Bounded storage: at most MAX_SESSIONS entries and a soft total byte cap.
 *   - Oversized sessions lose illustration previews before any other field.
 *   - Corrupt JSON in either store fails closed (returns empty / null)
 *     instead of throwing — the add-in must keep working with no sessions.
 *
 * Public surface (used by bootstrap + chat-view):
 *   listSessions, loadSession, saveSession, deleteSession, clearAllSessions,
 *   generateTitle.
 *
 * @module taskpane/sessions
 */

import { newId, normalizeMessage } from './message-shape.js';

const INDEX_KEY = 'wordAI.sessions.index';
const SESSION_KEY_PREFIX = 'wordAI.session.';
let activeScope = null;
const MAX_SESSIONS = 50;
const MAX_SESSION_BYTES = 1_500_000; // ~1.5 MB per session
const MAX_TOTAL_BYTES = 4_000_000;  // ~4 MB total across sessions + index

/**
 * Returns a stable session id.
 * @returns {string}
 */
function generateId() {
    return newId('s');
}

function nowIso() {
    return new Date().toISOString();
}

function scopePrefix() {
    if (!activeScope) return null;
    return `claric.history.${encodeURIComponent(JSON.stringify([activeScope.workspaceId, activeScope.documentId]))}.`;
}

function indexKey() {
    return activeScope ? `${scopePrefix()}index` : INDEX_KEY;
}

function sessionPrefix() {
    return activeScope ? `${scopePrefix()}session.` : SESSION_KEY_PREFIX;
}

function sessionKey(id) {
    return `${sessionPrefix()}${encodeURIComponent(id)}`;
}

export function setSessionScope(scope) {
    if (scope === null || scope === undefined) {
        activeScope = null;
        return null;
    }
    if (!scope || typeof scope.workspaceId !== 'string' || typeof scope.documentId !== 'string'
        || !scope.workspaceId || !scope.documentId) {
        throw new TypeError('setSessionScope: workspaceId and documentId are required');
    }
    activeScope = { workspaceId: scope.workspaceId, documentId: scope.documentId };
    return { ...activeScope };
}

export function getSessionScope() {
    return activeScope ? { ...activeScope } : null;
}

/**
 * v2 stores metadata and new session bodies together in one atomic write.
 * Entries without a body still point to their original v1 blob until updated.
 */
function readHistory() {
    const raw = localStorage.getItem(indexKey());
    let parsed;
    try { parsed = JSON.parse(raw); } catch { /* Ignore a corrupt index. */ }
    const current = parsed?.version === 2 && Array.isArray(parsed.entries);
    const entries = current ? parsed.entries : Array.isArray(parsed) ? parsed : [];
    return {
        current,
        entries: entries.filter((entry) => entry && typeof entry.id === 'string' && entry.id.trim())
            .map((entry) => ({ ...entry, updatedAt: typeof entry.updatedAt === 'string' ? entry.updatedAt : '' })),
    };
}

function commitHistory(entries, removedIds) {
    // setItem replaces the entire index or throws and leaves it intact.
    localStorage.setItem(indexKey(), JSON.stringify({ version: 2, entries }));
    for (const id of removedIds) {
        try { localStorage.removeItem(sessionKey(id)); } catch { /* Unreferenced legacy blob. */ }
    }
}

/**
 * Reads a single session by id. Returns null on missing/corrupt/wrong-shape
 * payloads (no message array → not a valid session).
 * @param {string} id
 * @returns {object|null}
 */
function readSession(id, history) {
    try {
        const entry = history.entries.find((item) => item.id === id);
        if (history.current && !entry) return null;
        const parsed = entry?.session ?? JSON.parse(localStorage.getItem(sessionKey(id)));
        if (!parsed || typeof parsed !== 'object') return null;
        if (!Array.isArray(parsed.messages)) return null;
        if (activeScope && (parsed.scope?.workspaceId !== activeScope.workspaceId || parsed.scope?.documentId !== activeScope.documentId)) return null;
        return parsed;
    } catch (_err) {
        return null;
    }
}

/**
 * Rough byte estimate via JSON.stringify length. Cheap and good enough for
 * our limits; we never need exact UTF-8 byte counts.
 * @param {object} value
 * @returns {number}
 */
function estimateBytes(value) {
    return JSON.stringify(value).length;
}

/** Per-item cap for proposal before/after diffs in persisted sessions. */
const MAX_PROPOSAL_DIFF_CHARS = 2_000;
/** Per-message cap for pathological chat text in persisted sessions. */
const MAX_MESSAGE_TEXT_CHARS = 100_000;

/**
 * Degrades a session until it fits under MAX_SESSION_BYTES, cheapest loss
 * first:
 *   1. illustration previews (previewSrc/previewSvg — pure eye-candy, regenerable),
 *   2. proposal before/after diff bodies (review metadata stays),
 *   3. whole proposals on a message,
 *   4. pathological message text (keeps the head; only hit by multi-MB
 *      single messages).
 * Mutates the passed object. Without stages 2-4 a large document run
 * overflowed the per-session cap and the whole session silently vanished
 * from history on QuotaExceededError.
 * @param {object} session
 */
function trimOversizedSession(session) {
    let bytes = estimateBytes(session);
    if (bytes <= MAX_SESSION_BYTES) return session;
    for (const msg of session.messages) {
        if (!msg || !Array.isArray(msg.proposals)) continue;
        for (const p of msg.proposals) {
            if (p && (p.previewSrc || p.previewSvg)) {
                bytes -= ((p.previewSrc || '').length + (p.previewSvg || '').length);
                p.previewSrc = null;
                p.previewSvg = null;
            }
        }
        if (bytes <= MAX_SESSION_BYTES) return session;
    }
    // Stage 2: truncate proposal diff bodies to reviewable metadata.
    for (const msg of session.messages) {
        if (!msg || !Array.isArray(msg.proposals)) continue;
        for (const p of msg.proposals) {
            if (!p || !Array.isArray(p.items)) continue;
            for (const item of p.items) {
                if (!item) continue;
                for (const field of ['before', 'after']) {
                    const value = item[field];
                    if (typeof value === 'string' && value.length > MAX_PROPOSAL_DIFF_CHARS) {
                        bytes -= value.length - MAX_PROPOSAL_DIFF_CHARS;
                        item[field] = value.slice(0, MAX_PROPOSAL_DIFF_CHARS);
                    }
                }
            }
        }
        if (bytes <= MAX_SESSION_BYTES) return session;
    }
    // Stage 3: drop whole proposals (oldest messages first).
    for (const msg of session.messages) {
        if (!msg || !Array.isArray(msg.proposals) || msg.proposals.length === 0) continue;
        bytes -= estimateBytes({ proposals: msg.proposals });
        msg.proposals = [];
        if (bytes <= MAX_SESSION_BYTES) return session;
    }
    // Stage 4: truncate pathological message text.
    for (const msg of session.messages) {
        if (!msg || typeof msg.text !== 'string' || msg.text.length <= MAX_MESSAGE_TEXT_CHARS) continue;
        bytes -= msg.text.length - MAX_MESSAGE_TEXT_CHARS;
        msg.text = msg.text.slice(0, MAX_MESSAGE_TEXT_CHARS);
        if (bytes <= MAX_SESSION_BYTES) return session;
    }
    return session;
}

/**
 * Picks a short preview snippet from the first user message. Returns ''
 * when no user message exists.
 * @param {Array<object>} messages
 * @returns {string}
 */
function makePreview(messages) {
    for (const m of messages) {
        if (m && m.role === 'user' && m.text && String(m.text).trim()) {
            return String(m.text).trim().replace(/\s+/g, ' ').slice(0, 80);
        }
    }
    return '';
}

/**
 * Removes unknown fields and normalizes empty values so persisted messages
 * never carry DOM references or transient closures. The shape (including the
 * attachment metadata reduction — image data URLs and extracted text would
 * blow the localStorage quota) lives in message-shape.js, shared with the
 * chat view's load path.
 * @param {object} m
 * @returns {object}
 */
function stripMessage(m) {
    const message = normalizeMessage(m);
    message.proposals = message.proposals.map((proposal) => proposal && ({
        ...proposal,
        ...(Array.isArray(proposal.items) ? { items: proposal.items.map((item) => item && { ...item }) } : {}),
    }));
    return message;
}

/**
 * Derives a default title from the first user message.
 * @param {Array<object>} messages
 * @returns {string}
 */
export function generateTitle(messages) {
    if (!Array.isArray(messages)) return 'Untitled chat';
    for (const m of messages) {
        if (m && m.role === 'user' && m.text && String(m.text).trim()) {
            const collapsed = String(m.text).trim().replace(/\s+/g, ' ');
            return collapsed.length > 30 ? collapsed.slice(0, 30) + '…' : collapsed;
        }
    }
    return 'Untitled chat';
}

/**
 * Lists session metadata, sorted by updatedAt descending (newest first).
 * Returns an empty array when no sessions exist or the index is corrupt.
 * @returns {Array<{id: string, title: string, createdAt: string, updatedAt: string, messageCount: number, preview: string}>}
 */
export function listSessions() {
    try {
        const { entries } = readHistory();
        return entries.map(({ session: _session, ...meta }) => meta)
            .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    } catch {
        return [];
    }
}

/**
 * Loads a single full session by id. Returns null when missing or corrupt.
 * @param {string} id
 * @returns {object|null}
 */
export function loadSession(id) {
    if (!id) return null;
    try { return readSession(id, readHistory()); } catch { return null; }
}

/**
 * Creates or updates a session. When opts.id matches an existing session,
 * title/createdAt are preserved and only messages/updatedAt are refreshed.
 *
 * @param {Array<object>} messages - The full message list for the session
 * @param {object} [opts]
 * @param {string} [opts.id] - Existing session id (omit to create new)
 * @param {string} [opts.title] - Explicit title override
 * @returns {object} The saved session (includes assigned id, timestamps)
 */
export function saveSession(messages, opts = {}) {
    if (!Array.isArray(messages)) {
        throw new TypeError('saveSession: messages must be an array');
    }

    const now = nowIso();
    const id = opts.id || generateId();
    const history = readHistory();
    const existing = opts.id ? readSession(opts.id, history) : null;
    const title = (opts.title && String(opts.title).trim())
        || (existing && existing.title)
        || generateTitle(messages);
    const createdAt = (existing && existing.createdAt) || now;

    const session = {
        id,
        title,
        createdAt,
        updatedAt: now,
        ...(activeScope ? { scope: { ...activeScope } } : {}),
        messages: messages.map(stripMessage),
    };
    trimOversizedSession(session);

    const meta = {
        id,
        title,
        createdAt,
        updatedAt: now,
        messageCount: session.messages.length,
        preview: makePreview(session.messages),
    };

    const entries = [{ ...meta, session }, ...history.entries.filter((entry) => entry.id !== id)];
    const removed = entries.splice(MAX_SESSIONS);
    // Include referenced legacy blobs in the same budget as embedded sessions.
    const entrySizes = entries.map((entry) => estimateBytes(entry) + 1
        + (entry.session ? 0 : (localStorage.getItem(sessionKey(entry.id)) || '').length));
    let totalBytes = estimateBytes({ version: 2, entries: [] }) + entrySizes.reduce((sum, size) => sum + size, 0);
    while (totalBytes > MAX_TOTAL_BYTES && entries.length > 1) {
        removed.push(entries.pop());
        totalBytes -= entrySizes.pop();
    }
    commitHistory(entries, new Set([id, ...removed.map((entry) => entry.id)]));
    return session;
}

/**
 * Deletes one session by id. Unknown ids are a no-op.
 * @param {string} id
 */
export function deleteSession(id) {
    if (!id) return;
    try {
        const { entries } = readHistory();
        commitHistory(entries.filter((entry) => entry.id !== id), [id]);
    } catch { /* Preserve the last committed history. */ }
}

/**
 * Wipes every persisted session and the index. Safe to call when empty.
 */
export function clearAllSessions() {
    try {
        for (let i = localStorage.length - 1; i >= 0; i--) {
            const key = localStorage.key(i);
            const prefix = sessionPrefix();
            if (key && key.startsWith(prefix)) {
                try { localStorage.removeItem(key); } catch (_err) { /* ignore */ }
            }
        }
    } catch (_err) {
        // Storage access can be disabled by the host.
    }
    try { localStorage.removeItem(indexKey()); } catch (_err) { /* ignore */ }
}

/**
 * Test seam. Not part of the public API; tests assert behavior through
 * the public functions but may read these constants.
 */
export const __testing = {
    INDEX_KEY,
    SESSION_KEY_PREFIX,
    MAX_SESSIONS,
    MAX_SESSION_BYTES,
    MAX_TOTAL_BYTES,
    generateId,
    nowIso,
    stripMessage,
    makePreview,
    trimOversizedSession,
};
