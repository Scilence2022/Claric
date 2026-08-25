/**
 * Sentence Diff Strategy for sentence-level diff application
 *
 * This strategy tokenizes by sentence boundaries to handle larger structural changes.
 * Falls back to block replacement if it fails.
 *
 * Vendored from office-word-diff (Apache-2.0) — see LICENSE and NOTICE in
 * this directory.
 *
 * @module lib/word-diff/sentence-diff
 */

import DiffMatchPatch from './diff-wordmode.js';
import { applyBlockReplaceStrategy } from './block-replace.js';

/**
 * Helper function for sentence-mode diff
 * Tokenizes text by sentence boundaries and computes diff
 *
 * @private
 * @param {string} text1 - Original text
 * @param {string} text2 - New text
 * @returns {Array<[number, string]>} Diff operations
 */
function diff_sentenceMode(text1, text2) {
    var dmp = new DiffMatchPatch();

    function tokenizeToSentences(text) {
        var sentences = [];
        var sentenceMap = {};
        var encoded = '';
        var remaining = text;

        while (remaining.length > 0) {
            var match1 = remaining.indexOf('. ');
            var match2 = remaining.indexOf('.  ');
            var nextBoundary = -1;
            var boundaryLen = 0;

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
                var sentence = remaining;
                if (sentence.length > 0) {
                    if (!Object.prototype.hasOwnProperty.call(sentenceMap, sentence)) {
                        sentences.push(sentence);
                        sentenceMap[sentence] = sentences.length - 1;
                    }
                    encoded += String.fromCharCode(sentenceMap[sentence]);
                }
                break;
            } else {
                // eslint-disable-next-line no-redeclare -- upstream redeclares this hoisted var intentionally
                var sentence = remaining.substring(0, nextBoundary + boundaryLen);
                remaining = remaining.substring(nextBoundary + boundaryLen);

                if (!Object.prototype.hasOwnProperty.call(sentenceMap, sentence)) {
                    sentences.push(sentence);
                    sentenceMap[sentence] = sentences.length - 1;
                }
                encoded += String.fromCharCode(sentenceMap[sentence]);
            }
        }
        return { encoded: encoded, sentences: sentences };
    }

    var result1 = tokenizeToSentences(text1);
    var result2 = tokenizeToSentences(text2);

    var sentenceArray = [''];
    var sentenceToIndex = {};

    [result1.sentences, result2.sentences].forEach(sentenceList => {
        sentenceList.forEach(sentence => {
            if (!Object.prototype.hasOwnProperty.call(sentenceToIndex, sentence)) {
                sentenceArray.push(sentence);
                sentenceToIndex[sentence] = sentenceArray.length - 1;
            }
        });
    });

    var chars1 = '';
    result1.sentences.forEach(sentence => {
        chars1 += String.fromCharCode(sentenceToIndex[sentence]);
    });

    var chars2 = '';
    result2.sentences.forEach(sentence => {
        chars2 += String.fromCharCode(sentenceToIndex[sentence]);
    });

    var diffs = dmp.diff_main(chars1, chars2, false);
    dmp.diff_charsToLines_(diffs, sentenceArray);

    return diffs;
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
 * @returns {Promise<{strategy: string, insertions: number, deletions: number}>}
 * @throws {Error} If all strategies fail
 */
export async function applySentenceDiffStrategy(context, range, text1, text2, log) {
    log("DEBUG: Running Sentence Diff Strategy...");

    let insertions = 0;
    let deletions = 0;
    let diffs;

    try {
        diffs = diff_sentenceMode(text1, text2);
    } catch (e) {
        log(`❌ Error in diff_sentenceMode: ${e.message}`);
        log("⚠️ Falling back to Block Replace Strategy...");
        return await applyBlockReplaceStrategy(context, range, text2, log);
    }

    log(`DEBUG: DMP generated ${diffs.length} chunks.`);

    try {
        // Enable Track Changes
        let trackingEnabled = false;
        try {
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
                await context.sync();
                trackingEnabled = true;
                log("DEBUG: Track changes enabled");
            }
        } catch (e) {
            log(`⚠️ Could not enable track changes: ${e.message}`);
        }

        // Strategy: Sentence Map
        // 1. Get ranges for sentences using sentence separators
        const sentenceRanges = range.getTextRanges([". ", ".  "], false);
        sentenceRanges.load("items/text");
        await context.sync();

        log(`DEBUG: Found ${sentenceRanges.items.length} sentence ranges.`);

        // 2. Build Sentence Map (Index -> Range)
        const sentenceMap = sentenceRanges.items.map((r, index) => ({ index, range: r, text: r.text }));

        // Pass 1: Deletions
        log("DEBUG: Pass 1 - Deletions");
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
            log(`DEBUG: Deleting ${deleteTargets.length} sentences...`);
            deleteTargets.reverse().forEach(t => {
                t.range.delete();
            });
            deletions = deleteTargets.length;
            await context.sync();
        }

        // Pass 2: Insertions
        log("DEBUG: Pass 2 - Insertions");
        const deletedIndices = new Set(deleteTargets.map(t => t.index));
        const sentencesAfterDeletes = sentenceMap.filter(t => !deletedIndices.has(t.index));

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
        log("✅ Sentence-level diff operations applied");

        if (trackingEnabled) {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            await context.sync();
        }

        return { strategy: 'sentence', insertions, deletions };

    } catch (e) {
        log(`❌ Sentence-level strategy failed: ${e.message}`);
        log("⚠️ Falling back to Block Replace Strategy...");

        // Reset range and fall back to block replacement
        range.insertText(text1, Word.InsertLocation.replace);
        await context.sync();

        return await applyBlockReplaceStrategy(context, range, text2, log);
    }
}
