/**
 * Chat message shape.
 *
 * The single definition of a persisted chat message. Both legs of the
 * save/load round-trip normalize through here — sessions.js on the way into
 * localStorage, ui/chat-view.js on the way back into the DOM — so a new
 * message field cannot be validated on one side and silently dropped on the
 * other.
 *
 * No DOM and no storage access: pure shape functions.
 *
 * @module taskpane/message-shape
 */

import { attachmentMeta } from '../lib/file-attachments.js';

/**
 * Returns a stable id with the given prefix ('s' for sessions, 'm' for
 * messages). Uses crypto.randomUUID when available (Office.js WebView2 does),
 * else a timestamp + random suffix.
 *
 * @param {string} prefix
 * @returns {string}
 */
export function newId(prefix) {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `${prefix}-${crypto.randomUUID()}`;
    }
    return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * Keeps only display metadata from an attachment list. Extracted text and
 * image data URLs never reach the session store (localStorage is ~5 MB) or
 * the restored DOM — the chip row and history only need name/kind/size.
 * Re-exported from the upload layer so the reduction has one definition.
 *
 * @param {Array<{name: string, kind: string, size: number}>} attachments
 * @returns {Array<{name: string, kind: string, size: number}>}
 */
export const normalizeAttachments = attachmentMeta;

/**
 * Normalizes one message into the canonical persisted shape, dropping
 * unknown fields and any DOM references or transient closures a live message
 * handle may have carried.
 *
 * @param {object} m - Live or stored message (non-objects yield defaults)
 * @returns {{id: string, role: 'user'|'assistant', text: string, status: string,
 *   error: string|null, worklog: {count: number, durationMs: number}|null,
 *   model: {sections: number}|null, citations: Array<object>,
 *   proposals: Array<object>, attachments: Array<object>, ts: string}}
 */
export function normalizeMessage(m) {
    const safe = m && typeof m === 'object' ? m : {};
    return {
        id: safe.id || newId('m'),
        role: safe.role === 'assistant' ? 'assistant' : 'user',
        text: typeof safe.text === 'string' ? safe.text : '',
        status: typeof safe.status === 'string' ? safe.status : '',
        error: typeof safe.error === 'string' ? safe.error : null,
        worklog: safe.worklog && typeof safe.worklog === 'object'
            ? { count: Number(safe.worklog.count) || 0, durationMs: Number(safe.worklog.durationMs) || 0 }
            : null,
        model: safe.model && typeof safe.model === 'object'
            ? { sections: Number(safe.model.sections) || 0 }
            : null,
        citations: normalizeCitations(safe.citations),
        proposals: Array.isArray(safe.proposals) ? safe.proposals : [],
        attachments: normalizeAttachments(safe.attachments),
        ts: typeof safe.ts === 'string' ? safe.ts : new Date().toISOString(),
    };
}

/**
 * Normalizes citation pills to the two fields the pill row renders.
 *
 * @param {Array<{label: string, searchText: string}>} citations
 * @returns {Array<{label: string, searchText: string}>}
 */
export function normalizeCitations(citations) {
    if (!Array.isArray(citations)) return [];
    return citations
        .filter((c) => c && typeof c === 'object')
        .map((c) => ({ label: String(c.label || ''), searchText: String(c.searchText || '') }));
}
