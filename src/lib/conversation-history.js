/**
 * Conversation history → model-message conversion with a token budget.
 *
 * Two pure functions, no DOM/storage/global state:
 *   - buildConversationHistory(records, { log, config }) turns persisted
 *     session records into a detached, role-preserving history array bounded
 *     by the configured token budget (whole turns, oldest dropped first).
 *   - withConversationHistory(messages, history) inserts an already-budgeted
 *     history array into an outgoing request after its leading system
 *     messages, without touching the current request.
 *
 * The budget is an estimate, not a tokenizer: CJK characters count ≈1 token
 * each, everything else ≈1 token per 4 characters. It is deliberately
 * conservative for Latin text and roughly accurate for CJK so a fixed budget
 * behaves similarly across languages.
 *
 * @module conversation-history
 */

/** Default chat-history budget, sized for modern 128K+-context backends. */
export const DEFAULT_HISTORY_BUDGET_TOKENS = 128_000;
export const MIN_HISTORY_BUDGET_TOKENS = 1_000;
export const MAX_HISTORY_BUDGET_TOKENS = 2_000_000;
/** Tokens held back from history for the current request plus model output. */
const REQUEST_RESERVE_TOKENS = 8_192;

const MAX_DETAIL_CHARS = 1000;
const MAX_PROPOSALS = 8;
const MAX_ITEMS = 8;
const MAX_ATTACHMENTS = 8;

/** CJK blocks (Hangul, kana, CJK unified/compatibility, full-width forms). */
const CJK_CHAR_RE = /[\u1100-\u11ff\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/;

/**
 * Estimates the token count of a string (see module docblock).
 * @param {unknown} value
 * @returns {number}
 */
export function estimateTokens(value) {
    const text = typeof value === 'string' ? value : '';
    if (!text) return 0;
    let cjkChars = 0;
    for (const ch of text) {
        if (CJK_CHAR_RE.test(ch)) cjkChars++;
    }
    return cjkChars + Math.ceil((text.length - cjkChars) / 4);
}

/**
 * Resolves the effective history budget from config. `config.contextBudgetTokens`
 * is the total conversation-context allowance; a fixed reserve is held back so
 * history cannot crowd out the current request or the model's output.
 *
 * @param {{contextBudgetTokens?: number}} [config]
 * @returns {number} Effective history budget in estimated tokens
 */
export function historyBudgetTokens(config) {
    const raw = config && Number.isFinite(config.contextBudgetTokens)
        ? config.contextBudgetTokens
        : DEFAULT_HISTORY_BUDGET_TOKENS;
    const budget = Math.min(
        Math.max(Math.round(raw), MIN_HISTORY_BUDGET_TOKENS),
        MAX_HISTORY_BUDGET_TOKENS,
    );
    return Math.max(MIN_HISTORY_BUDGET_TOKENS, budget - REQUEST_RESERVE_TOKENS);
}

function validMessage(message) {
    return message && (message.role === 'user' || message.role === 'assistant')
        && typeof message.content === 'string' && message.content.trim();
}

function turnTokens(turn) {
    return turn.reduce((sum, message) => sum + estimateTokens(message.content), 0);
}

/** Truncates content to the token budget, keeping the head, marking the cut. */
function fitToTokens(content, budgetTokens) {
    const total = estimateTokens(content);
    if (total <= budgetTokens) return content;
    let end = Math.max(1, Math.floor(content.length * (budgetTokens / total)));
    while (end > 1 && estimateTokens(content.slice(0, end)) > budgetTokens) {
        end = Math.max(1, end - Math.max(64, Math.floor(end / 8)));
    }
    return content.slice(0, end).trimEnd() + ' [trimmed]';
}

/**
 * Keeps a contiguous suffix of whole turns within the token budget. The most
 * recent turn is always kept — dropping it would strand the model without its
 * immediate context, the exact continuity bug this module exists to fix — so
 * an oversized latest turn is truncated (oldest message first kept whole,
 * newest truncated) instead of dropped.
 * @param {Array<{role: 'user'|'assistant', content: string}>} messages - Freshly built, mutable
 * @param {number} budgetTokens
 * @param {function(string): void} [onTrim]
 * @returns {Array<{role: 'user'|'assistant', content: string}>}
 */
function recentTurns(messages, budgetTokens, onTrim) {
    const turns = [];
    for (const message of messages) {
        if (message.role === 'user' || turns.length === 0) turns.push([]);
        turns[turns.length - 1].push(message);
    }
    if (turns.length === 0) return [];

    let start = turns.length - 1;
    if (turnTokens(turns[start]) > budgetTokens) {
        let remaining = budgetTokens;
        for (let i = turns[start].length - 1; i >= 0 && remaining > 0; i--) {
            const message = turns[start][i];
            const tokens = estimateTokens(message.content);
            if (tokens > remaining) {
                message.content = fitToTokens(message.content, remaining);
                remaining = 0;
            } else {
                remaining -= tokens;
            }
        }
        onTrim?.(`the latest turn exceeded the ${budgetTokens}-token history budget and was truncated`);
    }
    let size = turnTokens(turns[start]);
    while (start > 0) {
        const length = turnTokens(turns[start - 1]);
        if (size + length > budgetTokens) break;
        size += length;
        start--;
    }
    if (start > 0) onTrim?.(`${start} older turn(s) omitted to fit the ${budgetTokens}-token history budget`);
    return turns.slice(start).flat();
}

/**
 * Converts persisted session records into a bounded model-history array.
 *
 * Failed/cancelled assistant turns keep a truthful status line but drop their
 * partial output; proposal state is summarized so the model never mistakes
 * "proposed" for "applied". Attachment names survive as metadata with an
 * explicit note that the original bytes are gone.
 *
 * @param {unknown} records - Session message records
 * @param {{log?: Function, config?: {contextBudgetTokens?: number}}} [options]
 * @returns {Array<{role: 'user'|'assistant', content: string}>} Detached messages
 */
export function buildConversationHistory(records, { log, config } = {}) {
    const budgetTokens = historyBudgetTokens(config);
    const report = (detail) => {
        if (typeof log === 'function') log(`Conversation history trimmed: ${detail}.`, 'info');
    };
    let shortened = 0;
    const detail = (value) => {
        const valueText = typeof value === 'string' ? value : '';
        if (valueText.length <= MAX_DETAIL_CHARS) return valueText;
        shortened++;
        return `${valueText.slice(0, MAX_DETAIL_CHARS)} [trimmed]`;
    };
    const take = (values, limit) => {
        if (!Array.isArray(values)) return [];
        if (values.length > limit) shortened++;
        return values.slice(0, limit);
    };
    const messages = [];
    for (const record of Array.isArray(records) ? records : []) {
        if (!record || (record.role !== 'user' && record.role !== 'assistant')) continue;
        const parts = [];
        const status = typeof record.status === 'string' ? record.status : '';
        const cancelled = record.cancelled === true || /\bcancell?ed\b/i.test(status);
        if (record.role === 'user' || (!record.error && !cancelled)) {
            if (typeof record.text === 'string' && record.text.trim()) parts.push(record.text);
        }
        if (record.role === 'user') {
            for (const attachment of take(record.attachments, MAX_ATTACHMENTS)) {
                if (!attachment || !(typeof attachment.name === 'string' && attachment.name.trim())) continue;
                const kind = typeof attachment.kind === 'string' ? attachment.kind : '';
                parts.push(`Attachment: ${detail(attachment.name)}${kind ? ` (${detail(kind)})` : ''}. Original file contents/bytes are not available in this history; reattach the file if needed.`);
            }
        } else {
            if (cancelled) parts.push('Turn cancelled; incomplete assistant output omitted.');
            if (record.error) parts.push(`Turn failed: ${detail(record.error) || 'error'}. Incomplete assistant output omitted.`);
            if (status.trim()) parts.push(`Status: ${detail(status)}`);
            for (const proposal of take(record.proposals, MAX_PROPOSALS)) {
                if (!proposal || typeof proposal !== 'object') continue;
                const states = {
                    pending: 'pending; proposed only, not applied',
                    applied: 'applied; selected changes were applied, not necessarily every listed item',
                    rejected: 'rejected; not applied',
                    warning: 'warning; application may be partial, verify the document',
                    error: 'error; application failed or may be partial, verify the document',
                };
                const state = Object.hasOwn(states, proposal.state) ? states[proposal.state] : 'unknown; application not confirmed';
                parts.push(`Proposal: ${detail(proposal.title) || 'Untitled'} [${state}]`);
                if (typeof proposal.detail === 'string' && proposal.detail.trim()) parts.push(`Proposal detail: ${detail(proposal.detail)}`);
                if (typeof proposal.countsText === 'string' && proposal.countsText.trim()) parts.push(`Proposal counts: ${detail(proposal.countsText)}`);
                for (const item of take(proposal.items, MAX_ITEMS)) {
                    if (!item || typeof item !== 'object') continue;
                    parts.push(`Proposed item: ${detail(item.label)}\nBefore: ${detail(item.before)}\nProposed after: ${detail(item.after)}`);
                }
            }
        }
        if (parts.length) messages.push({ role: record.role, content: parts.join('\n\n') });
    }
    if (shortened) report(`${shortened} metadata/diff field(s) or list(s) shortened`);
    return recentTurns(messages, budgetTokens, report);
}

/**
 * Inserts a pre-budgeted history array into an outgoing request: leading
 * trusted system messages stay first, history next, current request last.
 * Budget enforcement belongs to buildConversationHistory (which owns config);
 * this function only validates and orders. Copies history objects; never
 * mutates either input; no global memory.
 *
 * @param {Array<{role: string, content: string|Array<object>}>} messages
 * @param {Array<{role: string, content: string}>} [history]
 * @returns {Array} New request array
 */
export function withConversationHistory(messages, history = []) {
    const request = Array.isArray(messages) ? messages : [];
    let leadingSystems = 0;
    while (request[leadingSystems]?.role === 'system') leadingSystems++;
    const previous = (Array.isArray(history) ? history : [])
        .filter(validMessage)
        .map(({ role, content }) => ({ role, content }));
    return [...request.slice(0, leadingSystems), ...previous, ...request.slice(leadingSystems)];
}
