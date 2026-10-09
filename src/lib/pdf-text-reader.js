/**
 * PDF.js text streams use getReader for Safari/WKWebView compatibility.
 * @param {{streamTextContent: Function}} page
 * @param {{signal?: AbortSignal, maxChars?: number, separator?: string,
 *   preserveEOL?: boolean, onLimit?: 'truncate'|'error'}} [options]
 * @returns {Promise<{text: string, truncated: boolean}>}
 */
export async function readPdfPageText(page, { signal, maxChars = 30000, separator = ' ',
    preserveEOL = false, onLimit = 'truncate' } = {}) {
    if (!Number.isInteger(maxChars) || maxChars < 0 || !['truncate', 'error'].includes(onLimit)) {
        throw new Error('Invalid PDF text limit.');
    }
    const aborted = () => new DOMException('PDF text reading cancelled.', 'AbortError');
    if (signal?.aborted) throw aborted();
    const reader = page.streamTextContent().getReader();
    let finished = false;
    let cancelRequested = false;
    let rejectAbort;
    const cancellation = new Promise((_resolve, reject) => { rejectAbort = reject; });
    const cancel = (reason) => {
        if (cancelRequested) return;
        cancelRequested = true;
        try { void Promise.resolve(reader.cancel?.(reason)).catch(() => {}); }
        catch (_error) { /* Cleanup must not replace a text result or failure. */ }
    };
    const abort = () => { const error = aborted(); cancel(error); rejectAbort(error); };
    signal?.addEventListener('abort', abort, { once: true });
    let text = '';
    let seenItem = false;
    try {
        // A reader API does not require ReadableStream[Symbol.asyncIterator].
        while (true) {
            if (signal?.aborted) throw aborted();
            const { done, value } = await Promise.race([reader.read(), cancellation]);
            if (signal?.aborted) throw aborted();
            if (done) { finished = true; return { text, truncated: false }; }
            if (!Array.isArray(value?.items)) throw new Error('PDF text stream returned invalid items.');
            for (const item of value.items) {
                if (typeof item?.str !== 'string') continue;
                const piece = (seenItem ? separator : '') + item.str + (preserveEOL && item.hasEOL ? '\n' : '');
                seenItem = true;
                const remaining = maxChars - text.length;
                if (piece.length > remaining) {
                    if (onLimit === 'error') throw new Error(`Extracted PDF text exceeds the ${maxChars.toLocaleString('en-US')} character limit.`);
                    return { text: text + piece.slice(0, remaining), truncated: true };
                }
                text += piece;
            }
        }
    } finally {
        signal?.removeEventListener('abort', abort);
        if (!finished) cancel();
        try { reader.releaseLock(); }
        catch (_error) { /* Cleanup must not replace a text result or failure. */ }
    }
}
