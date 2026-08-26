/**
 * Sentence Diff Strategy for sentence-level diff application
 *
 * This strategy tokenizes by sentence boundaries to handle larger structural changes.
 * Falls back to block replacement if it fails.
 *
 * Vendored from office-word-diff (Apache-2.0) — see LICENSE and NOTICE in
 * this directory.
 *
 * Local modifications:
 *   - diff_sentenceMode diffs the occurrence-ordered sentence sequence.
 *     Upstream built the diff input from the DEDUPED sentence list, so any
 *     repeated sentence silently shifted every later sentence out of
 *     alignment (wrong deletions/insertions with no error raised).
 *   - Accepts options.trackChanges (default true); when false the strategy
 *     does not touch the document's changeTrackingMode (caller owns it).
 *   - Typed activity-log messages instead of DEBUG/emoji lines.
 *
 * @module lib/word-diff/sentence-diff
 */

import DiffMatchPatch from './diff-wordmode.js';
import { applyBlockReplaceStrategy } from './block-replace.js';

/**
 * Tokenizes text into sentences in OCCURRENCE order. A sentence boundary is
 * '. ' or '.  ' (period + space(s)); the boundary stays attached to the
 * sentence it terminates. Repeated sentences keep their positions — the diff
 * input must mirror the document's actual sentence sequence.
 *
 * @private
 * @param {string} text
 * @returns {string[]} Sentences in document order (deduped nowhere)
 */
function tokenizeToSentences(text) {
    const sentences = [];
    let remaining = text;

    while (remaining.length > 0) {
        const match1 = remaining.indexOf('. ');
        const match2 = remaining.indexOf('.  ');
        let nextBoundary = -1;
        let boundaryLen = 0;

        if (match1 !== -1 && match2 !== -1) {
            if (match1 < match2) {
                nextBoundary = match1;
                boundaryLen = 2;
            } else {
                nextBoundary = match2;
                boundaryLen = 3;
            }
        } else if (match1 !== -1) {
            nextBoundary = match1;
            boundaryLen = 2;
        } else if (match2 !== -1) {
            nextBoundary = match2;
            boundaryLen = 3;
        }

        if (nextBoundary === -1) {
            sentences.push(remaining);
            break;
        }
        sentences.push(remaining.substring(0, nextBoundary + boundaryLen));
        remaining = remaining.substring(nextBoundary + boundaryLen);
    }
    return sentences;
}

/**
 * Helper function for sentence-mode diff.
 * Encodes each sentence as one char (via a shared dedupe table across both
 * texts) and runs diff-match-patch on the encoded sequences.
 * Pure function — exported for tests.
 *
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @returns {Array<[number, string]>} Diff operations
 */
export function diff_sentenceMode(text1, text2) {
    const dmp = new DiffMatchPatch();

    const sentences1 = tokenizeToSentences(text1);
    const sentences2 = tokenizeToSentences(text2);

    const sentenceArray = [''];
    const sentenceToIndex = {};
    for (const sentenceList of [sentences1, sentences2]) {
        for (const sentence of sentenceList) {
            if (!Object.prototype.hasOwnProperty.call(sentenceToIndex, sentence)) {
                sentenceArray.push(sentence);
                sentenceToIndex[sentence] = sentenceArray.length - 1;
            }
        }
    }

    const encode = (sentences) => sentences
        .map((sentence) => String.fromCharCode(sentenceToIndex[sentence]))
        .join('');

    const diffs = dmp.diff_main(encode(sentences1), encode(sentences2), false);
    // Deliberately reuses dmp's private charsToLines to decode the sentence-encoded sequences.
    dmp.diff_charsToLines_(diffs, sentenceArray);

    return /** @type {Array<[number, string]>} */ (diffs);
}

/**
 * Applies the "Sentence Diff" strategy.
 * Tokenizes by sentence boundaries to handle larger structural changes.
 * Falls back to block replacement if it fails.
 *
 * @param {Word.RequestContext} context - The Word request context
 * @param {Word.Range} range - The target range to update
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @param {function} log - Callback for logging messages
 * @param {object} [options]
 * @param {boolean} [options.trackChanges=true] - When false the strategy does
 *   NOT touch the document's changeTrackingMode; the caller owns it.
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 * @throws {Error} If all strategies fail
 */
export async function applySentenceDiffStrategy(context, range, text1, text2, log, options = {}) {
    const trackChanges = options.trackChanges !== false;
    log('Running sentence-level diff...', 'info');

    let insertions = 0;
    let deletions = 0;
    let diffs;

    try {
        diffs = diff_sentenceMode(text1, text2);
    } catch (e) {
        log(`Sentence diff computation failed (${e.message}); falling back to block replace`, 'warning');
        return applyBlockReplaceStrategy(context, range, text2, log, options);
    }

    // True only when THIS strategy enabled tracking (and must restore it).
    let trackingEnabled = false;

    try {
        if (trackChanges && Word.ChangeTrackingMode) {
            try {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
                await context.sync();
                trackingEnabled = true;
            } catch (e) {
                log(`Could not enable track changes: ${e.message}`, 'warning');
            }
        }

        // Strategy: Sentence Map
        // 1. Get ranges for sentences using sentence separators
        const sentenceRanges = range.getTextRanges(['. ', '.  '], false);
        sentenceRanges.load('items/text');
        await context.sync();

        // 2. Build Sentence Map (Index -> Range)
        const sentenceMap = sentenceRanges.items.map((r, index) => ({ index, range: r, text: r.text }));

        // Pass 1: Deletions
        let sentenceIndex = 0;
        const deleteTargets = [];

        for (const [op, _chunk] of diffs) {
            if (op === 0) { // EQUAL
                sentenceIndex++;
            } else if (op === -1) { // DELETE
                if (sentenceIndex < sentenceMap.length) {
                    deleteTargets.push(sentenceMap[sentenceIndex]);
                    sentenceIndex++;
                }
            }
        }

        if (deleteTargets.length > 0) {
            deleteTargets.reverse().forEach((t) => {
                t.range.delete();
            });
            deletions = deleteTargets.length;
            await context.sync();
        }

        // Pass 2: Insertions
        const deletedIndices = new Set(deleteTargets.map((t) => t.index));
        const sentencesAfterDeletes = sentenceMap.filter((t) => !deletedIndices.has(t.index));

        let currentSentenceIdx = 0;
        let lastAnchorRange = null;

        for (const [op, chunk] of diffs) {
            if (op === 0) { // EQUAL
                if (currentSentenceIdx < sentencesAfterDeletes.length) {
                    lastAnchorRange = sentencesAfterDeletes[currentSentenceIdx].range;
                    currentSentenceIdx++;
                }
            } else if (op === 1) { // INSERT
                if (lastAnchorRange) {
                    lastAnchorRange.insertText(chunk, Word.InsertLocation.after);
                } else {
                    // If no anchor (start of text), insert at start of range
                    range.getRange(Word.RangeLocation.start).insertText(chunk, Word.InsertLocation.before);
                }
                insertions++;
            }
        }

        await context.sync();
        log(`Sentence-level diff applied (${insertions} insertions, ${deletions} deletions)`, 'info');

        return { strategy: 'sentence', insertions, deletions };
    } catch (e) {
        log(`Sentence-level strategy failed (${e.message}); falling back to block replace`, 'warning');

        // Reset the range with tracking OFF so the reset itself does not show
        // up as a spurious whole-range revision pair, then block-replace
        // (which manages tracking itself via the same options).
        if (trackingEnabled) {
            try {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
                trackingEnabled = false;
            } catch (_resetErr) {
                // Best-effort; block replace re-asserts tracking itself.
            }
        }
        range.insertText(text1, Word.InsertLocation.replace);
        await context.sync();

        return applyBlockReplaceStrategy(context, range, text2, log, options);
    } finally {
        if (trackingEnabled) {
            try {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                await context.sync();
            } catch (_restoreErr) {
                // Best-effort restore; never mask the primary outcome.
            }
        }
    }
}
