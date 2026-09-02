/**
 * Sessions storage layer.
 *
 * Persists chat sessions to localStorage under two key families:
 *   - wordAI.sessions.index : metadata list (id, title, createdAt,
 *                             updatedAt, messageCount, preview)
 *   - wordAI.session.<id>   : full session payload incl. messages array
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

function sessionKey(id) {
    return `${SESSION_KEY_PREFIX}${id}`;
}

/**
 * Reads the index array. Always returns an array (empty on missing/corrupt).
 * @returns {Array<object>}
 */
function readIndex() {
    try {
        const raw = localStorage.getItem(INDEX_KEY);
        if (!raw) return [];
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed : [];
    } catch (_err) {
        return [];
    }
}

function writeIndex(idx) {
    try {
        localStorage.setItem(INDEX_KEY, JSON.stringify(idx));
        return true;
    } catch (_err) {
        // Quota exceeded — the caller decides whether this is fatal.
        return false;
    }
}

/**
 * Reads a single session by id. Returns null on missing/corrupt/wrong-shape
 * payloads (no message array → not a valid session).
 * @param {string} id
 * @returns {object|null}
 */
function readSession(id) {
    try {
        const raw = localStorage.getItem(sessionKey(id));
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== 'object') return null;
        if (!Array.isArray(parsed.messages)) return null;
        return parsed;
    } catch (_err) {
        return null;
    }
}

/**
 * Best-effort write; throws only when localStorage is fundamentally broken.
 * @param {object} session
 */
function writeSession(session) {
    localStorage.setItem(sessionKey(session.id), JSON.stringify(session));
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
    return normalizeMessage(m);
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
    const idx = readIndex();
    return [...idx].sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}

/**
 * Loads a single full session by id. Returns null when missing or corrupt.
 * @param {string} id
 * @returns {object|null}
 */
export function loadSession(id) {
    if (!id) return null;
    return readSession(id);
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
    const existing = opts.id ? readSession(opts.id) : null;
    const title = (opts.title && String(opts.title).trim())
        || (existing && existing.title)
        || generateTitle(messages);
    const createdAt = (existing && existing.createdAt) || now;

    const session = {
        id,
        title,
        createdAt,
        updatedAt: now,
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

    const idx = readIndex();
    const existingIdx = idx.findIndex((m) => m.id === id);
    if (existingIdx >= 0) {
        idx.splice(existingIdx, 1);
    }
    idx.unshift(meta);

    // Cap by count: drop the oldest metadata entries (and their blobs).
    while (idx.length > MAX_SESSIONS) {
        const dropped = idx.pop();
        try { localStorage.removeItem(sessionKey(dropped.id)); } catch (_err) { /* ignore */ }
    }

    // Total-size cap: drop oldest non-current sessions until the index +
    // every remaining blob fits under MAX_TOTAL_BYTES. Eviction runs BEFORE
    // the new blob is written: writing first could blow the localStorage
    // quota while stale blobs were still holding the very space the eviction
    // was about to free — the current session was lost in exactly the case
    // the cap was meant to protect against. Reads raw stored string lengths
    // instead of JSON.parsing each blob — estimateBytes is a stringify-length
    // count, so the result is identical without re-serializing up to
    // MAX_SESSIONS blobs on every committed turn.
    let totalBytes = estimateBytes(idx);
    for (const m of idx) {
        if (m.id === id) {
            totalBytes += estimateBytes(session);
        } else {
            const raw = localStorage.getItem(sessionKey(m.id));
            if (raw) totalBytes += raw.length;
        }
    }
    // Walk from oldest (tail) toward newest, dropping until under cap.
    while (totalBytes > MAX_TOTAL_BYTES && idx.length > 1) {
        const dropped = idx.pop();
        if (dropped.id === id) {
            // Current session alone is over the cap — keep it; the per-session
            // trim above already cut the largest field. Bail out.
            idx.push(dropped);
            break;
        }
        const raw = localStorage.getItem(sessionKey(dropped.id));
        const bytes = raw ? raw.length : 0;
        try { localStorage.removeItem(sessionKey(dropped.id)); } catch (_err) { /* ignore */ }
        totalBytes -= bytes + estimateBytes(dropped);
    }

    // The cap math has now made room when it could. If the write STILL fails
    // (encoding overhead, other-origin inflation), the last resort is
    // dropping every other session rather than losing the current one.
    try {
        writeSession(session);
    } catch (_err) {
        for (const m of idx) {
            if (m.id !== id) {
                try { localStorage.removeItem(sessionKey(m.id)); } catch (_err2) { /* ignore */ }
            }
        }
        writeSession(session);
    }

    if (!writeIndex(idx)) {
        // The blob is saved but the history index is now stale/inconsistent.
        // Surface it — the caller logs to the activity log — instead of
        // silently dropping the failure on the floor.
        throw new Error('Session saved, but the history index could not be updated (storage quota). History may be stale.');
    }
    return session;
}

/**
 * Deletes one session by id. Unknown ids are a no-op.
 * @param {string} id
 */
export function deleteSession(id) {
    if (!id) return;
    try { localStorage.removeItem(sessionKey(id)); } catch (_err) { /* ignore */ }
    const idx = readIndex().filter((m) => m.id !== id);
    writeIndex(idx);
}

/**
 * Wipes every persisted session and the index. Safe to call when empty.
 */
export function clearAllSessions() {
    const idx = readIndex();
    for (const m of idx) {
        try { localStorage.removeItem(sessionKey(m.id)); } catch (_err) { /* ignore */ }
    }
    try { localStorage.removeItem(INDEX_KEY); } catch (_err) { /* ignore */ }
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
