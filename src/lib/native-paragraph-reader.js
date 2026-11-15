import { queueRevisionRead, resolveRevisionRead, RevisionSafetyError, wordErrorDetail } from './word-revisions.js';

export const NATIVE_PARAGRAPH_BATCH_SIZE = 24;

const paragraphOptions = { paragraph: true, singleParagraphContent: true };

function checkCancelled(signal) {
    if (signal?.aborted) throw new DOMException('Document reading cancelled.', 'AbortError');
}

function paragraphReadError(first, last, cause) {
    const target = first === last ? `paragraph ${first}` : `paragraphs ${first}–${last}`;
    const error = new Error(`Word could not read ${target}: ${wordErrorDetail(cause)}`, { cause });
    return Object.assign(error, { code: cause?.code, debugInfo: cause?.debugInfo });
}

/**
 * Read revision states while the caller owns the native paragraph proxies.
 * @param {Word.RequestContext} context
 * @param {Word.Paragraph[]} paragraphs
 * @param {{properties?: string, includeTable?: boolean, indexes?: number[], signal?: AbortSignal}} [options]
 */
export async function readNativeParagraphs(context, paragraphs, {
    properties = 'text', includeTable = false, indexes, signal,
} = {}) {
    const paragraphNumber = (offset) => (indexes?.[offset] ?? offset) + 1;
    async function readBatch(start, end, includeReviewedText = true) {
        checkCancelled(signal);
        const batch = paragraphs.slice(start, end).map((paragraph) => {
            paragraph.load(properties);
            const read = queueRevisionRead(paragraph, { ...paragraphOptions, includeReviewedText });
            const table = includeTable ? paragraph.parentTableOrNullObject : null;
            table?.load('isNullObject');
            return { paragraph, read, table };
        });
        try {
            await context.sync();
        } catch (error) {
            checkCancelled(signal);
            if (error?.code === 'GeneralException' || error?.message === 'GeneralException') {
                if (batch.length > 1) {
                    const middle = start + Math.floor(batch.length / 2);
                    const left = await readBatch(start, middle);
                    return left.concat(await readBatch(middle, end));
                }
                // Reviewed text helps scope ambiguous exports. When its
                // bridge call fails, the paragraph export still supplies the
                // revision baseline and normal XML scoping/protection rules.
                if (batch[0].read?.currentText) return readBatch(start, end, false);
            }
            throw paragraphReadError(paragraphNumber(start), paragraphNumber(end - 1), error);
        }
        checkCancelled(signal);
        return batch.map(({ paragraph, read, table }, index) => {
            try {
                const revisionState = resolveRevisionRead(paragraph, read, paragraphOptions);
                if (!includeReviewedText && (revisionState.paragraphCount !== 1 || !revisionState.fingerprint)) {
                    throw new RevisionSafetyError('Word did not return XML scoped to one paragraph.');
                }
                return { revisionState, inTable: !!table && !table.isNullObject };
            } catch (error) {
                const number = paragraphNumber(start + index);
                throw paragraphReadError(number, number, error);
            }
        });
    }

    const records = [];
    for (let start = 0; start < paragraphs.length; start += NATIVE_PARAGRAPH_BATCH_SIZE) {
        records.push(...await readBatch(start, Math.min(start + NATIVE_PARAGRAPH_BATCH_SIZE, paragraphs.length)));
    }
    return records;
}
