import { parseDocument } from '../lib/document-parser.js';
import { queueRevisionRead, resolveRevisionRead } from '../lib/word-revisions.js';

const secretKey = /(?:api.?key|authorization|password|secret|^(?:access|refresh|pairing)?_?token$|^headers$)/i;
const binaryKey = /^(?:dataUrl|base64|blob|imageData)$/i;

/** Redact configured credentials wherever they occur, including chat/log text. */
export function redactContext(text, config = {}) {
    const secrets = [];
    function visit(value, sensitive = false) {
        if (typeof value === 'string' && sensitive && value) secrets.push(value);
        else if (value && typeof value === 'object') {
            Object.entries(value).forEach(([key, item]) => visit(item, sensitive || secretKey.test(key)));
        }
    }
    visit(config);
    let safe = String(text);
    for (const secret of secrets.sort((a, b) => b.length - a.length)) {
        for (const value of new Set([secret, encodeURIComponent(secret), JSON.stringify(secret).slice(1, -1)])) {
            safe = safe.split(value).join('[redacted]');
        }
    }
    return safe.replace(/(Bearer\s+)[^\s"'<>]+/gi, '$1[redacted]')
        .replace(/([?&](?:api[_-]?key|token|access_token|key|secret|password)=)[^&#\s"'<>]+/gi, '$1[redacted]')
        .replace(/(https?:\/\/)[^\s/@]+:[^\s/@]+@/gi, '$1[redacted]@');
}

function serializable(value) {
    return JSON.stringify(value, (key, item) => secretKey.test(key) ? '[redacted]'
        : binaryKey.test(key) || (typeof item === 'string' && /^(?:data:|blob:)/.test(item))
            ? '[binary attachment omitted]' : item, 2);
}

export async function buildContextReport({ state, session, chatText, logs, draft, attachments, metadata }) {
    const parts = ['Claric context report', `Captured at: ${new Date().toISOString()}`,
        'Includes current document and chat content. Credentials are redacted; binary attachments are omitted.',
        '\n## Environment and settings', serializable({ ...metadata, platform: state.platform,
            processing: state.isProcessing || state.isProcessingDoc || state.isProcessingSummary,
            config: state.config, prompts: state.promptManager?.getState() }),
        '\n## Current chat (including model activity and proposals)', chatText || '(empty)',
        '\n## Chat session data', serializable(session), '\n## Pending input', draft || '(empty)',
        '\n## Attached files', attachments || '(none)'];
    // These reads never select, modify, accept or reject document content.
    try {
        let selection;
        await Word.run(async (context) => {
            const range = context.document.getSelection();
            range.load('text');
            const read = queueRevisionRead(range);
            await context.sync();
            selection = resolveRevisionRead(range, read).text;
        });
        parts.push('\n## Current selection', selection || '(empty)');
    } catch (error) { parts.push('\n## Current selection', `Unavailable: ${error.message}`); }
    try {
        const model = await parseDocument();
        parts.push('\n## Current document (current revision view)', model.paragraphs.map((p) => p.text).join('\n'));
    } catch (error) { parts.push('\n## Current document', `Unavailable: ${error.message}`); }
    parts.push('\n## Activity log', logs || '(empty)');
    return redactContext(parts.join('\n'), state.config);
}
