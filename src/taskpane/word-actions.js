/**
 * Word Actions Module
 *
 * The document/LLM operations behind chat turns, adapted from the original
 * taskpane.js handlers. Differences from the old handlers:
 *
 * - Handlers take explicit args ({ category, promptTemplate, commentInstructions })
 *   instead of reading the PromptManager active prompt and DOM fields.
 * - Selection-scope amendment runs are split into prepare (LLM call, returns
 *   the proposed edit) and apply (diff into Word as tracked changes), so the
 *   chat UI can stage a proposal card between the two.
 * - Document-scope runs take an onProgress callback and return citation data
 *   instead of writing to a fixed progress bar.
 *
 * Dependencies flow in via a `deps` object: { appState, log, logWithRetry,
 * updateStatusBar }. Module-level code is side-effect free; Word/Office
 * globals are only touched inside functions.
 *
 * @module word-actions
 */

import { applyTokenMapStrategy, applySentenceDiffStrategy } from 'office-word-diff';
import { hasCjk, applyCharDiffStrategy } from '../lib/char-diff.js';
import { sendPrompt, sendPromptStream, stripMarkdown } from '../lib/llm-client.js';
import { fireCommentRequest } from '../lib/comment-request.js';
import { extractAllComments, extractDocumentStructured, estimateTokenCount, extractTrackedChanges, extractCommentsOnRange } from '../lib/comment-extractor.js';
import { formatSelectionWithComments } from '../lib/selection-with-comments.js';
import { createSummaryDocument, buildSummaryHtml } from '../lib/document-generator.js';
import { parseDelimitedResponse, buildFallbackClassificationPrompt } from '../lib/response-parser.js';
import { parseDocument } from '../lib/document-parser.js';
import { chunkDocument } from '../lib/document-chunker.js';
import { extractContext } from '../lib/context-extractor.js';
import { processChunksParallel } from '../lib/orchestrator.js';
import { bookmarkChunkRanges, applyChunkResults, cleanupBookmarks } from '../lib/reassembler.js';
import { getActiveBackendConfig } from './app-state.js';

/**
 * Flattens a chat-completions messages array into a single prompt string
 * (system content first), matching the old sendPrompt compatibility shim.
 *
 * @param {Array<{role: string, content: string}>} messages
 * @returns {string}
 * @private
 */
function _flattenMessages(messages) {
    if (messages.length >= 2 && messages[0].role === 'system') {
        return messages[0].content + '\n\n' + messages.slice(1).map(m => m.content).join('\n\n');
    }
    return messages.map(m => m.content).join('\n\n');
}

/**
 * Builds a PromptManager-compatible shim that serves `template` as the active
 * prompt for `category` while delegating everything else (context prompt,
 * other categories) to the real promptManager. Used to run the existing
 * orchestrator/comment pipelines against an explicit skill template.
 *
 * @param {object} promptManager - Real PromptManager instance
 * @param {string} category - Pipeline category the template belongs to
 * @param {string} template - Explicit prompt template
 * @returns {object} PromptManager-compatible shim
 */
export function makePromptShim(promptManager, category, template) {
    return {
        getActivePrompt(cat) {
            if (cat === category) {
                return { id: `skill-${category}`, name: 'Skill template', template };
            }
            return promptManager.getActivePrompt(cat);
        },
        getActiveMode() {
            return category;
        },
        composeMessages: (selectionText, cat) =>
            cat === category
                ? promptManager.composeMessages(selectionText, cat, template)
                : promptManager.composeMessages(selectionText, cat),
    };
}

/**
 * Reads the current selection, optionally enriched with anchored comments
 * (config.includeCommentsInSelection). Throws when the selection is empty.
 *
 * @param {object} deps - { appState, log }
 * @returns {Promise<{ selectionText: string, plainSelectionText: string }>}
 */
export async function readSelectionText(deps) {
    const { appState, log } = deps;
    const includeComments = !!appState.config.includeCommentsInSelection;
    let selectionText = '';
    let plainSelectionText = '';
    let enrichmentError = null;

    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        // OOXML fetch only when enrichment is requested (toggle ON) — saves a sync round-trip on the default path.
        const ooxmlResult = includeComments ? selection.getOoxml() : null;
        await context.sync();
        if (!selection.text || !selection.text.trim()) {
            throw new Error('Please select some text first.');
        }
        plainSelectionText = selection.text;

        if (!includeComments) {
            selectionText = plainSelectionText;
            return;
        }

        let comments = [];
        try {
            comments = await extractCommentsOnRange(context, selection);
        } catch (err) {
            console.error('[readSelectionText] extractCommentsOnRange failed', { err });
            enrichmentError = err;
        }

        if (!enrichmentError) {
            try {
                selectionText = formatSelectionWithComments(ooxmlResult.value, comments);
            } catch (err) {
                console.error('[readSelectionText] formatSelectionWithComments failed', { err });
                enrichmentError = err;
            }
        }
    });

    if (includeComments && enrichmentError) {
        log(`Comment enrichment failed (${enrichmentError.message}); falling back to plain selection.`, 'warning');
        selectionText = plainSelectionText;
    } else if (includeComments && selectionText.length > plainSelectionText.length) {
        log(`Selection enriched with comment threads (+${selectionText.length - plainSelectionText.length} chars)`, 'info');
    }

    return { selectionText, plainSelectionText };
}

/**
 * Returns whether the document currently has a non-collapsed selection.
 * Resolves to false when the Word API is unavailable.
 *
 * @returns {Promise<boolean>}
 */
export async function hasNonEmptySelection() {
    try {
        let has = false;
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            has = !!(selection.text && selection.text.trim());
        });
        return has;
    } catch (_err) {
        return false;
    }
}

/**
 * Runs the selection-scope amendment LLM call WITHOUT applying the result.
 * The returned proposal is staged in the chat UI; applySelectionAmendment
 * writes it to the document.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.promptTemplate - Amendment instruction/template ({selection} placeholder supported)
 * @param {string} [args.commentInstructions] - When non-empty, merged amendment + comment mode
 * @returns {Promise<{ selectionText: string, amendedText: string|null, commentText: string|null, model: string }>}
 */
export async function prepareSelectionAmendment(deps, { promptTemplate, commentInstructions } = {}) {
    const { appState, log } = deps;
    const { selectionText } = await readSelectionText(deps);
    const backendConfig = getActiveBackendConfig(appState);
    log(`Processing selection (${selectionText.length} chars) via ${backendConfig.model}...`, 'info');

    const merged = !!(commentInstructions && commentInstructions.trim());
    const messages = merged
        ? appState.promptManager.composeMergedMessages(selectionText, commentInstructions, promptTemplate)
        : appState.promptManager.composeMessages(selectionText, 'amendment', promptTemplate);

    if (messages.length === 0) {
        throw new Error('No prompt composed — check the skill template');
    }

    const rawResponse = await sendPrompt(backendConfig, _flattenMessages(messages), log);
    log(`LLM response received [${backendConfig.model}]`, 'success');

    if (!merged) {
        return {
            selectionText,
            amendedText: stripMarkdown(rawResponse, log),
            commentText: null,
            model: backendConfig.model,
        };
    }

    // Merged mode: parse the ===AMENDMENT=== / ===COMMENT=== protocol,
    // with the fallback classification call preserved from the old handler.
    let parsed = parseDelimitedResponse(rawResponse);

    if (parsed.amendment === null) {
        log('Response missing delimiters, attempting to classify...', 'info');
        const fallbackMessages = buildFallbackClassificationPrompt(rawResponse, selectionText);
        try {
            const fallbackResponse = await sendPrompt(backendConfig, _flattenMessages(fallbackMessages), log);
            parsed = parseDelimitedResponse(fallbackResponse);
            if (parsed.amendment === null) {
                log('Could not split response into amendment and comment', 'warning');
                parsed = { amendment: rawResponse.trim(), comment: null, raw: rawResponse };
            }
        } catch (fallbackError) {
            log(`Fallback classification failed: ${fallbackError.message}`, 'warning');
            parsed = { amendment: rawResponse.trim(), comment: null, raw: rawResponse };
        }
    }

    return {
        selectionText,
        amendedText: parsed.amendment ? stripMarkdown(parsed.amendment, log) : null,
        commentText: parsed.comment || null,
        model: backendConfig.model,
    };
}

/**
 * Applies a prepared selection amendment to the document as tracked changes
 * (per config.trackChangesEnabled), then inserts the optional comment.
 *
 * @param {object} deps - { appState, log }
 * @param {object} proposal - Result of prepareSelectionAmendment
 */
export async function applySelectionAmendment(deps, proposal) {
    const { appState, log } = deps;
    const { selectionText, amendedText, commentText } = proposal;

    if (amendedText) {
        log('Applying changes...', 'info');
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            if (Word.ChangeTrackingMode) {
                context.document.changeTrackingMode = appState.config.trackChangesEnabled
                    ? Word.ChangeTrackingMode.trackAll
                    : Word.ChangeTrackingMode.off;
            }
            if (appState.config.lineDiffEnabled) {
                await applySentenceDiffStrategy(context, selection, selectionText, amendedText, log);
            } else if (hasCjk(selectionText) || hasCjk(amendedText)) {
                // CJK text has no word boundaries for the token map — use
                // char-level diff so e.g. a one-comma edit stays minimal.
                await applyCharDiffStrategy(context, selection, selectionText, amendedText, log);
            } else {
                await applyTokenMapStrategy(context, selection, selectionText, amendedText, log);
            }
        });
        log('Changes applied successfully', 'success');
    }

    if (commentText && appState.supportsComments) {
        log('Inserting comment...', 'info');
        try {
            await Word.run(async (context) => {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                const contentRange = selection.getRange();
                contentRange.insertComment(commentText);
                await context.sync();
            });
            log('Comment inserted successfully', 'success');
        } catch (commentError) {
            // Comment insertion failed -- log the comment text so it is not lost
            log(`Comment insertion failed: ${commentError.message}. Comment text: "${commentText}"`, 'warning');
        }
    } else if (commentText && !appState.supportsComments) {
        log(`Comment generated but Word API 1.4 not available. Comment: "${commentText}"`, 'warning');
    }
}

/**
 * Fires the fire-and-forget comment pipeline on the current selection
 * (bookmark capture + async LLM call + comment insertion with retry link).
 *
 * @param {object} deps - { appState, log, logWithRetry, updateStatusBar }
 * @param {object} args
 * @param {string} args.promptTemplate - Explicit comment template
 */
export async function fireSelectionComment(deps, { promptTemplate } = {}) {
    const { appState, log, logWithRetry, updateStatusBar } = deps;
    const { selectionText } = await readSelectionText(deps);
    fireCommentRequest(selectionText, {
        config: getActiveBackendConfig(appState),
        sendPromptFn: sendPrompt,
        promptManager: makePromptShim(appState.promptManager, 'comment', promptTemplate),
        commentQueue: appState.commentQueue,
        log,
        addLogWithRetryFn: logWithRetry,
        updateStatusBarFn: updateStatusBar,
    });
}

/**
 * Runs the whole-document chunked pipeline (parse → chunk → context →
 * parallel LLM → apply as tracked changes/comments). Preserves the old
 * handleProcessDocument semantics; cancellation via appState.processDocController.
 *
 * @param {object} deps - { appState, log, logWithRetry }
 * @param {object} args
 * @param {string} args.category - 'amendment' | 'comment'
 * @param {string} args.promptTemplate - Explicit template for the pipeline
 * @param {string} [args.commentInstructions] - Merged-mode comment instructions
 * @param {function} [args.onProgress] - Progress callback from the orchestrator
 * @param {boolean} [args.gateApply=false] - When true, stop after the LLM
 *   processing phase and return an apply/discard continuation instead of
 *   writing to the document (used to stage a proposal card for amendments).
 * @returns {Promise<{ results: Array, applicationResult?: object, chunks: Array, cancelled?: boolean,
 *   staged?: boolean, apply?: Function, discard?: Function, failedCount?: number, cancelledCount?: number }>}
 */
export async function runDocumentSkill(deps, { category, promptTemplate, commentInstructions = '', onProgress, gateApply = false } = {}) {
    const { appState, log, logWithRetry } = deps;
    const signal = appState.processDocController ? appState.processDocController.signal : undefined;
    const promptShim = makePromptShim(appState.promptManager, category, promptTemplate);

    // Step 1: Parse document
    log('Parsing document...', 'info');
    const docModel = await parseDocument();
    log(`Found ${docModel.paragraphs.length} paragraphs (~${docModel.totalTokens} tokens)`, 'info');

    // Step 2: Chunk document
    const chunks = chunkDocument(docModel, { maxTokens: 6000 });
    log(`Split into ${chunks.length} chunks`, 'info');

    // Step 3: Extract context
    const documentContext = extractContext(docModel);
    log(`Extracted ${documentContext.definitions.length} definitions, ${documentContext.outline.length} headings`, 'info');

    // Step 4: Bookmark chunk ranges
    const bookmarkMap = await bookmarkChunkRanges(chunks);

    // Step 5: Process chunks in parallel
    const backendConfig = getActiveBackendConfig(appState);
    const concurrency = chunks.some(c => c.tokenCount > 8000) ? 4 : 6;

    const results = await processChunksParallel(chunks, {
        config: backendConfig,
        promptManager: promptShim,
        documentContext: documentContext,
        log,
        onProgress,
        signal,
        concurrency,
        timeoutMs: 300000,
        commentInstructions,
    });

    const failed = results.filter(r => r.status === 'rejected').length;
    const cancelled = results.filter(r => r.status === 'cancelled').length;

    /**
     * Applies the processed results to the document (tracked changes +
     * comments), cleans up bookmarks, logs the summary, and registers the
     * retry link for failed chunks.
     * @returns {Promise<object>} applicationResult from applyChunkResults
     */
    const apply = async () => {
        log('Applying changes to document...', 'info');
        const applicationResult = await applyChunkResults(results, bookmarkMap, {
            trackChangesEnabled: appState.config.trackChangesEnabled,
            lineDiffEnabled: appState.config.lineDiffEnabled,
            log,
            commentGranularity: appState.config.commentGranularity,
        });

        await cleanupBookmarks(bookmarkMap);

        log(
            `Document processed: ${chunks.length} chunks, ` +
            `${applicationResult.amendmentsApplied} amendments applied, ` +
            `${applicationResult.commentsInserted} comments inserted` +
            (failed > 0 ? `, ${failed} chunks failed` : '') +
            (cancelled > 0 ? `, ${cancelled} chunks cancelled` : ''),
            failed > 0 ? 'warning' : 'success'
        );

        if (failed > 0 && logWithRetry) {
            const failedChunks = results.filter(r => r.status === 'rejected');
            logWithRetry(
                `${failed} chunk(s) failed. Click to retry failed chunks.`,
                'warning',
                () => retryFailedChunks(deps, { failedResults: failedChunks, bookmarkMap, backendConfig, promptShim, onProgress })
            );
        }

        return applicationResult;
    };

    /**
     * Discards a staged run: removes the hidden chunk bookmarks without
     * touching document text.
     */
    const discard = async () => {
        await cleanupBookmarks(bookmarkMap);
        log('Proposed changes discarded; no edits were applied.', 'info');
    };

    // Gated mode (amendment pipeline): stop before writing to the document
    // and hand the caller an apply/discard continuation so the chat UI can
    // stage a proposal card for user confirmation.
    if (gateApply) {
        return { staged: true, results, chunks, apply, discard, failedCount: failed, cancelledCount: cancelled };
    }

    const applicationResult = await apply();
    return { results, applicationResult, chunks, cancelled: cancelled > 0 };
}

/**
 * Retries processing only the failed chunks of a document-scope run.
 * Re-runs the orchestrator on the failed chunk subset and applies results.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {Array} args.failedResults - ChunkResult objects with status 'rejected'
 * @param {Map} args.bookmarkMap - Original chunkId -> bookmarkName map
 * @param {object} args.backendConfig - Backend configuration
 * @param {object} args.promptShim - PromptManager shim from the original run
 * @param {function} [args.onProgress] - Progress callback
 */
export async function retryFailedChunks(deps, { failedResults, bookmarkMap, backendConfig, promptShim, onProgress } = {}) {
    const { appState, log } = deps;

    if (appState.isProcessingDoc) {
        log('Document processing is already running. Wait for it to finish before retrying.', 'warning');
        return;
    }

    log(`Retrying ${failedResults.length} failed chunk(s)...`, 'info');
    appState.isProcessingDoc = true;
    appState.processDocController = new AbortController();

    try {
        const retryChunks = failedResults.map(r => ({
            id: r.chunkId,
            text: r.originalText || '',
            tokenCount: r.originalText ? Math.ceil(r.originalText.length / 4) : 0,
            overlapText: '',
        }));

        const results = await processChunksParallel(retryChunks, {
            config: backendConfig,
            promptManager: promptShim,
            documentContext: null,
            log,
            onProgress,
            signal: appState.processDocController.signal,
            concurrency: 4,
            timeoutMs: 300000,
            commentInstructions: '',
        });

        const applicationResult = await applyChunkResults(results, bookmarkMap, {
            trackChangesEnabled: appState.config.trackChangesEnabled,
            lineDiffEnabled: appState.config.lineDiffEnabled,
            log,
            commentGranularity: appState.config.commentGranularity,
        });

        const stillFailed = results.filter(r => r.status === 'rejected').length;
        log(
            `Retry complete: ${applicationResult.amendmentsApplied} amendments, ` +
            `${applicationResult.commentsInserted} comments` +
            (stillFailed > 0 ? `, ${stillFailed} still failed` : ''),
            stillFailed > 0 ? 'warning' : 'success'
        );
    } catch (error) {
        if (error.name === 'AbortError') {
            log('Retry cancelled.', 'warning');
        } else {
            log(`Retry failed: ${error.message}`, 'error');
        }
    } finally {
        appState.isProcessingDoc = false;
        appState.processDocController = null;
    }
}

/**
 * Runs the summary pipeline: extract comments (+ optional document text and
 * tracked changes per template placeholders), send to LLM, open a new Word
 * document with the formatted summary.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.promptTemplate - Summary template; supports {comments},
 *   {whole document}, {tracked changes} placeholders
 */
export async function runSummarySkill(deps, { promptTemplate } = {}) {
    const { appState, log } = deps;

    log('Extracting document comments...', 'info');
    const comments = await extractAllComments();
    if (comments.length === 0) {
        log('No review comments found — summarizing from document text only.', 'info');
    }

    const summaryOpts = { templateOverride: promptTemplate };

    if (promptTemplate.includes('{whole document}')) {
        const richness = (appState.config.docExtraction || {}).richness || 'structured';
        log(`Extracting document text (${richness})...`, 'info');
        summaryOpts.documentText = await extractDocumentStructured({ richness });
        log(`Document text extracted (${summaryOpts.documentText.length} chars, ~${estimateTokenCount(summaryOpts.documentText)} tokens)`, 'info');
    }

    if (appState.config.trackedChangesExtraction && promptTemplate.includes('{tracked changes}')) {
        log('Extracting tracked changes (OOXML parsing)...', 'info');
        const tcResult = await extractTrackedChanges();
        log(`Tracked changes extracted (${tcResult.changes.length} change(s))`, 'info');

        let tcText = '';
        if (tcResult.changes.length > 0) {
            tcText = tcResult.changes.map((c, i) => {
                const num = i + 1;
                const author = c.author || 'Unknown';
                const date = c.date || '';
                const dateStr = date ? ` on ${date}` : '';

                if (c.type === 'Replaced') {
                    return `[Change ${num}] REPLACED by ${author}${dateStr}:\n` +
                           `  BEFORE: "${c.beforeText}"\n` +
                           `  AFTER:  "${c.afterText}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                } else if (c.type === 'Deleted') {
                    return `[Change ${num}] DELETED by ${author}${dateStr}:\n` +
                           `  REMOVED: "${c.text}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                } else if (c.type === 'Added') {
                    return `[Change ${num}] ADDED by ${author}${dateStr}:\n` +
                           `  INSERTED: "${c.text}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                } else if (c.type.startsWith('Moved')) {
                    return `[Change ${num}] ${c.type.toUpperCase()} by ${author}${dateStr}:\n` +
                           `  TEXT: "${c.text}"` +
                           (c.paragraphText ? `\n  IN CLAUSE: "${c.paragraphText}"` : '');
                }
                return `[Change ${num}] ${c.type} by ${author}${dateStr}: "${c.text}"`;
            }).join('\n\n');
        }

        if (tcText) {
            summaryOpts.trackedChangesText = tcText;
        } else if (tcResult.changes.length === 0) {
            summaryOpts.trackedChangesText = '(No tracked changes found in document)';
        }
    }

    const messages = appState.promptManager.composeSummaryMessages(comments, summaryOpts);
    if (messages.length === 0) {
        throw new Error('No summary prompt composed — check the skill template');
    }

    const backendConfig = getActiveBackendConfig(appState);
    log(`Sending summary request [${backendConfig.model}]...`, 'info');
    const llmResponse = await sendPrompt(backendConfig, _flattenMessages(messages), log);
    log(`Summary received (${llmResponse.length} chars). Creating document...`, 'info');

    let docTitle = 'Document Summary';
    try {
        await Word.run(async (context) => {
            const props = context.document.properties;
            props.load('title');
            await context.sync();
            if (props.title) {
                docTitle = `Summary - ${props.title}`;
            }
        });
    } catch (_titleErr) {
        // Title lookup failed -- use default
    }

    const html = buildSummaryHtml(llmResponse, comments, docTitle);
    await createSummaryDocument(html, docTitle, log);
    log('Summary document opened successfully.', 'success');

    return { chars: llmResponse.length, commentCount: comments.length };
}

/**
 * Answers a free-text question in chat using the document as context.
 * Streams tokens via sendPromptStream when the backend supports SSE.
 *
 * @param {object} deps - { appState, log }
 * @param {object} args
 * @param {string} args.question - The user's question
 * @param {string} [args.skillTemplate] - Persona/instruction template from a chat skill
 * @param {function} [args.onToken] - Called with each streamed token
 * @param {function} [args.onStatus] - Called with stage updates ("Reading the document...", "Waiting for model...")
 * @param {AbortSignal} [args.signal] - Cancellation signal
 * @returns {Promise<string>} The full answer text
 */
export async function answerQuestion(deps, { question, skillTemplate, onToken, onStatus, signal } = {}) {
    const { appState, log } = deps;

    const richness = (appState.config.docExtraction || {}).richness || 'structured';
    log('Extracting document text for context...', 'info');
    if (onStatus) onStatus('Reading the document...');
    const documentText = await extractDocumentStructured({ richness });

    let prompt = '';
    const contextPrompt = appState.promptManager.getActivePrompt('context');
    if (contextPrompt) {
        prompt += contextPrompt.template + '\n\n';
    }
    if (skillTemplate) {
        prompt += skillTemplate + '\n\n';
    }
    prompt += question + '\n\n--- DOCUMENT ---\n' + documentText;

    const backendConfig = getActiveBackendConfig(appState);
    log(`Asking [${backendConfig.model}]...`, 'info');
    if (onStatus) onStatus(`Waiting for ${backendConfig.model}...`);
    // Long documents + slow backends can exceed the 120s client default;
    // chat answers get 5 minutes (the doc pipeline uses the same per-chunk).
    return sendPromptStream(backendConfig, prompt, onToken, log, signal, 300000);
}

/**
 * Scrolls the Word document to the first occurrence of a text snippet.
 * Used by citation pills on completed document-scope runs.
 *
 * @param {object} deps - { log }
 * @param {string} searchText - Distinctive text near the start of the chunk
 * @returns {Promise<boolean>} True when a match was found and selected
 */
export async function revealTextSnippet(deps, searchText) {
    const { log } = deps;
    const needle = (searchText || '').trim().slice(0, 200);
    if (!needle) return false;

    try {
        let found = false;
        await Word.run(async (context) => {
            const results = context.document.body.search(needle, { matchCase: false, matchWholeWord: false });
            results.load('items');
            await context.sync();
            if (results.items.length > 0) {
                results.items[0].select();
                await context.sync();
                found = true;
            }
        });
        if (!found) {
            log('Could not locate the chunk text in the document (it may have changed).', 'warning');
        }
        return found;
    } catch (error) {
        log(`Jump to section failed: ${error.message}`, 'warning');
        return false;
    }
}
