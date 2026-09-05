/**
 * Orchestrator Module
 *
 * Parallel LLM dispatch engine for whole-document processing. Takes document
 * chunks (from document-chunker.js) and document context (from context-extractor.js),
 * composes per-chunk LLM prompts, dispatches them in parallel with a concurrency pool,
 * and returns results with status tracking.
 *
 * Key behaviors:
 * - Worker-pool concurrency pattern with configurable limit (default 4)
 * - Promise.allSettled semantics: failed chunks don't block successful ones
 * - AbortController cancellation stops pending work immediately
 * - Progress callback fires after each chunk with accurate counts and ETA
 * - Prompt composition includes document context prefix and overlap markers
 *
 * Pure JavaScript -- no Word API dependency.
 *
 * @module orchestrator
 */

import { sendMessages as defaultSendMessages, sendMessagesStream as defaultSendMessagesStream, stripMarkdown, stripChunkDelimiters } from './llm-client.js';
import { formatContextPrefix as defaultFormatContextPrefix } from './context-extractor.js';
import { withConversationHistory } from './conversation-history.js';
import {
  parseDelimitedResponse as defaultParseDelimitedResponse,
  defangProtocolMarkers,
  restoreProtocolMarkers,
} from './response-parser.js';

/**
 * @typedef {import('./document-chunker.js').DocumentChunk} DocumentChunk
 * @typedef {import('./context-extractor.js').DocumentContext} DocumentContext
 */

/**
 * @typedef {Object} ChunkResult
 * @property {string} chunkId - Matches DocumentChunk.id
 * @property {number} chunkIndex - Position in chunks array
 * @property {'fulfilled'|'rejected'|'cancelled'} status
 * @property {string|null} amendment - Amended text (for amendment/merged mode)
 * @property {string|null} comment - Comment text (for comment/merged mode)
 * @property {string|null} error - Error message if rejected
 * @property {string|null} reasoning - Model chain-of-thought when streamed
 * @property {DocumentChunk} chunk - Reference to original chunk
 */

/**
 * @typedef {Object} ProcessingProgress
 * @property {number} completed - Chunks successfully processed
 * @property {number} failed - Chunks that errored
 * @property {number} cancelled - Chunks cancelled by user
 * @property {number} total - Total chunks
 * @property {number} percentComplete - 0-100
 * @property {number} estimatedSecondsRemaining - ETA based on average per-chunk time
 */

/**
 * Composes the messages array for a single chunk's LLM call.
 *
 * @param {DocumentChunk} chunk
 * @param {DocumentContext} documentContext
 * @param {Object} promptManager
 * @param {string} mode - 'amendment'|'comment'|'both'
 * @param {string} commentInstructions - For merged mode
 * @param {function} formatContextPrefixFn
 * @returns {Array<{role: string, content: string}>}
 * @private
 */
function _composeChunkMessages(chunk, documentContext, promptManager, mode, commentInstructions, formatContextPrefixFn) {
  const messages = [];

  // Build chunk text from paragraphs
  const chunkText = chunk.paragraphs.map((p) => p.text).join('\n');

  // 1. System message: user's Context prompt (if active) + document context prefix
  const contextPrompt = promptManager.getActivePrompt('context');
  // A null/undefined context (e.g. retry runs that skip re-extraction)
  // contributes no prefix instead of crashing the prefix formatter.
  const docContextPrefix = documentContext
    ? formatContextPrefixFn(documentContext, chunkText, 4000)
    : '';

  let systemContent = '';
  if (contextPrompt) {
    systemContent += contextPrompt.template;
  }
  if (docContextPrefix) {
    if (systemContent) systemContent += '\n\n';
    systemContent += docContextPrefix;
  }
  if (systemContent) {
    messages.push({ role: 'system', content: systemContent });
  }

  // 2. User message with overlap markers and chunk text
  let userContent = '';

  // Get the appropriate prompt template based on mode
  let promptTemplate = '';
  if (mode === 'amendment' || mode === 'both') {
    const amendPrompt = promptManager.getActivePrompt('amendment');
    if (amendPrompt) promptTemplate = amendPrompt.template;
  } else if (mode === 'comment') {
    const commentPrompt = promptManager.getActivePrompt('comment');
    if (commentPrompt) promptTemplate = commentPrompt.template;
  }

  // Build the text content with overlap markers.
  //
  // Both the chunk body and the overlap context are defanged first: they are
  // untrusted document text being interpolated between framing markers, and a
  // document that reproduces `[END TEXT]` or `===COMMENT===` could otherwise
  // close the framing early (making the rest of the document read as
  // instructions) or make the response parser treat body text as a comment to
  // insert. The defang only perturbs text that literally reproduces a protocol
  // marker; in that rare case the returned amendment may differ from the
  // original by one zero-width character at that spot, which is a cosmetic
  // diff artifact and strictly preferable to a framing break.
  let textContent = '';
  if (chunk.overlapBefore) {
    textContent += `[CONTEXT - DO NOT AMEND]\n${defangProtocolMarkers(chunk.overlapBefore)}\n[END CONTEXT]\n\n`;
  }
  textContent += `[AMEND THIS TEXT]\n${defangProtocolMarkers(chunkText)}\n[END TEXT]`;

  // Substitute into template
  if (promptTemplate.includes('{selection}')) {
    userContent = promptTemplate.replace(/{selection}/g, textContent);
  } else {
    userContent = promptTemplate + '\n\n' + textContent;
  }

  // Add output format constraints for amendment mode
  if (mode === 'amendment' || mode === 'both') {
    userContent += `\n\nCRITICAL OUTPUT RULES:
- Output ONLY the amended text. Do not include any commentary, explanations, notes, summaries, or descriptions of your changes.
- Do NOT use markdown formatting. Output plain text only — no asterisks (*), no bold (**), no headings (###), no bullet points, no numbered lists unless they were in the original text.
- Preserve the original text structure. Only change content as instructed, not formatting.
- Do NOT add any preamble like "Here is the amended text:" or similar.
- Do NOT add any postscript explaining what was changed.
- Do NOT include the delimiter markers [AMEND THIS TEXT], [END TEXT], [CONTEXT - DO NOT AMEND], or [END CONTEXT] in your output. These are input framing only.`;
  }

  // For merged mode, append comment instructions with delimiter format.
  // When mode is 'amendment' but commentInstructions are provided, treat as merged.
  if ((mode === 'both' || mode === 'amendment') && commentInstructions) {
    userContent += `\n\nAdditionally, provide a comment for this text based on these instructions: ${commentInstructions.trim()}

FORMAT YOUR RESPONSE WITH THESE EXACT DELIMITERS:
===AMENDMENT===
[Your amended version of the text here]
===COMMENT===
[Your comment here]`;
  }

  messages.push({ role: 'user', content: userContent });

  return messages;
}

/**
 * Processes document chunks in parallel through the LLM with concurrency control.
 *
 * @param {DocumentChunk[]} chunks
 * @param {Object} options
 * @param {Object} options.config - LLM backend config { url, apiKey, model }
 * @param {Array<{role: string, content: string}>} [options.conversationHistory=[]] - Prior turns for every chunk
 * @param {Object} options.promptManager - PromptManager instance
 * @param {DocumentContext} options.documentContext - From extractContext()
 * @param {function} options.log - addLog callback
 * @param {function} [options.onProgress] - Called after each chunk with ProcessingProgress
 * @param {AbortSignal} [options.signal] - Cancellation signal
 * @param {number} [options.concurrency=4] - Max parallel LLM calls
 * @param {number} [options.timeoutMs=30000] - Per-chunk LLM timeout
 * @param {string} [options.commentInstructions=''] - Comment instructions for merged mode
 * @param {function} [options.onChunkToken] - When provided, chunks are sent with
 *   streaming enabled and this callback fires as (chunkInfo, kind, token) for
 *   each delta; chunkInfo is { id, index }, kind is 'content' or 'reasoning'
 * @param {function} [options.sendMessagesFn] - Injectable sendMessages (for testing)
 * @param {function} [options.sendMessagesStreamFn] - Injectable sendMessagesStream (for testing)
 * @param {function} [options.formatContextPrefixFn] - Injectable formatContextPrefix (for testing)
 * @param {function} [options.parseDelimitedResponseFn] - Injectable parseDelimitedResponse (for testing)
 * @returns {Promise<ChunkResult[]>}
 */
export async function processChunksParallel(chunks, options) {
  const {
    config,
    promptManager,
    documentContext,
    log,
    onProgress,
    signal,
    concurrency = 4,
    timeoutMs = 30000,
    commentInstructions = '',
    onChunkToken,
    sendMessagesFn = defaultSendMessages,
    sendMessagesStreamFn = defaultSendMessagesStream,
    formatContextPrefixFn = defaultFormatContextPrefix,
    parseDelimitedResponseFn = defaultParseDelimitedResponse,
  } = options;

  if (chunks.length === 0) {
    return [];
  }

  const mode = promptManager.getActiveMode();
  const results = new Array(chunks.length);
  let nextIndex = 0;
  let completed = 0;
  let failed = 0;
  let cancelled = 0;
  const chunkTimings = []; // Track per-chunk elapsed times for ETA

  function reportProgress() {
    if (!onProgress) return;

    const settled = completed + failed + cancelled;
    const remaining = chunks.length - settled;
    const percentComplete = Math.round((settled / chunks.length) * 100);

    // Estimate remaining time based on average chunk duration
    let estimatedSecondsRemaining = 0;
    if (remaining > 0 && chunkTimings.length > 0) {
      const avgMs = chunkTimings.reduce((a, b) => a + b, 0) / chunkTimings.length;
      estimatedSecondsRemaining = Math.round((remaining * avgMs) / 1000);
    }

    onProgress({
      completed,
      failed,
      cancelled,
      total: chunks.length,
      percentComplete,
      estimatedSecondsRemaining,
    });
  }

  function makeResult(chunkIndex, chunk, status, data = {}) {
    return {
      chunkId: chunk.id,
      chunkIndex,
      status,
      amendment: data.amendment || null,
      comment: data.comment || null,
      error: data.error || null,
      reasoning: data.reasoning || null,
      chunk,
    };
  }

  async function processChunk(chunkIndex) {
    const chunk = chunks[chunkIndex];
    const chunkStart = Date.now();

    // Check for cancellation before starting
    if (signal && signal.aborted) {
      cancelled++;
      results[chunkIndex] = makeResult(chunkIndex, chunk, 'cancelled');
      reportProgress();
      return;
    }

    try {
      // A single paragraph larger than the whole chunk budget cannot be split
      // (the reassembler bookmarks ranges by paragraph index and verifies
      // boundary text, so a half-paragraph chunk has no valid anchor). Sending
      // it anyway costs a round trip and comes back as an opaque backend error
      // about context length; failing here names the real cause and puts the
      // chunk on the retry path with the rest.
      if (chunk.oversized) {
        throw new Error(
          `Chunk ${chunk.id} holds a single paragraph of ~${chunk.tokenCount} tokens, over the `
          + 'per-chunk budget. Split that paragraph in the document, or amend it via a selection.'
        );
      }

      // Compose messages for this chunk
      const messages = withConversationHistory(_composeChunkMessages(
        chunk,
        documentContext,
        promptManager,
        mode,
        commentInstructions,
        formatContextPrefixFn
      ), options.conversationHistory);

      // Send to LLM. When the caller wants live tokens (chat UI model
      // activity view), use the streaming transport; otherwise the plain
      // request/response path keeps working for tests and headless callers.
      let responseText;
      let reasoningText = null;
      if (typeof onChunkToken === 'function') {
        const chunkInfo = { id: chunk.id, index: chunkIndex };
        const streamed = await sendMessagesStreamFn(config, messages, {
          onContent: (t) => onChunkToken(chunkInfo, 'content', t),
          onReasoning: (t) => onChunkToken(chunkInfo, 'reasoning', t),
        }, log, signal, timeoutMs);
        responseText = streamed.content;
        reasoningText = streamed.reasoning || null;
      } else {
        responseText = await sendMessagesFn(config, messages, log, signal, timeoutMs);
      }

      // Parse response based on mode.
      // When mode is 'amendment' but commentInstructions are provided,
      // the prompt requested delimited output -- parse it as merged.
      let amendment = null;
      let comment = null;
      const isMerged = (mode === 'both') || (mode === 'amendment' && commentInstructions);

      if (isMerged) {
        // Parse BEFORE restoring echoed markers. A defanged marker copied from
        // the document body must remain inert while section boundaries are
        // located; restoring first would let document text manufacture a new
        // amendment/comment delimiter.
        const parsed = parseDelimitedResponseFn(responseText);
        amendment = parsed.amendment;
        comment = parsed.comment;
        // No delimiters at all: the model ignored the output contract, so the
        // text's role is unknown. Treating it as the amendment (the old
        // behavior) fed prose like "I suggest the following changes: ..." into
        // the alignment pass as replacement text — the 30%-length truncation
        // guard cannot catch that, because the wrong text is the right length.
        // Reject instead so the chunk lands on the retry path with a reason.
        if (!amendment && !comment) {
          throw new Error(
            'Model response contained neither ===AMENDMENT=== nor ===COMMENT=== delimiters; ' +
            'cannot tell an amendment from commentary, so the chunk was not applied.'
          );
        }
      } else if (mode === 'amendment') {
        amendment = responseText;
      } else if (mode === 'comment') {
        comment = responseText;
      }

      // An empty response is recorded as fulfilled with no amendment (the
      // caller treats it as "nothing to change"), but it is not the same
      // thing as the model deliberately proposing no changes — surface it.
      if (!responseText || responseText.trim() === '') {
        log(`Chunk ${chunk.id}: LLM returned an empty response (may indicate a backend or model issue)`, 'warning');
      }

      // Post-process amendment text before restoring the zero-width character
      // used to defang echoed document markers. Restoring only after parsing
      // (and after cleanup) keeps a copied marker inert at every protocol
      // boundary while still returning the user's original visible text.
      if (amendment) {
        amendment = stripMarkdown(amendment, log);
        amendment = stripChunkDelimiters(amendment, log);
        amendment = restoreProtocolMarkers(amendment);
      }
      if (comment) {
        comment = restoreProtocolMarkers(comment);
      }

      completed++;
      chunkTimings.push(Date.now() - chunkStart);
      results[chunkIndex] = makeResult(chunkIndex, chunk, 'fulfilled', { amendment, comment, reasoning: reasoningText });
    } catch (error) {
      if (error.name === 'AbortError') {
        cancelled++;
        results[chunkIndex] = makeResult(chunkIndex, chunk, 'cancelled');
      } else if (error.name === 'TimeoutError') {
        failed++;
        chunkTimings.push(Date.now() - chunkStart);
        results[chunkIndex] = makeResult(chunkIndex, chunk, 'rejected', {
          error: error.message || String(error),
        });
        log(`Chunk ${chunk.id}: ${error.message}`, 'warning');
      } else {
        failed++;
        chunkTimings.push(Date.now() - chunkStart);
        const message = error.message || String(error);
        results[chunkIndex] = makeResult(chunkIndex, chunk, 'rejected', {
          error: message,
        });
        // Log at failure time: when every chunk fails there is no apply()
        // to carry the failure to the user, so this line is often the only
        // immediate trace of the backend problem.
        log(`Chunk ${chunk.id}: ${message}`, 'warning');
      }
    }

    reportProgress();
  }

  // Worker-pool pattern: spawn N workers, each pulls from shared index
  async function worker() {
    while (nextIndex < chunks.length) {
      // Check for cancellation before grabbing next chunk
      if (signal && signal.aborted) {
        // Mark all remaining unprocessed chunks as cancelled
        while (nextIndex < chunks.length) {
          const i = nextIndex++;
          if (!results[i]) {
            cancelled++;
            results[i] = makeResult(i, chunks[i], 'cancelled');
            reportProgress();
          }
        }
        return;
      }

      const i = nextIndex++;
      await processChunk(i);
    }
  }

  // Spawn concurrency workers
  const workerCount = Math.min(concurrency, chunks.length);
  const workers = Array.from({ length: workerCount }, () => worker());
  await Promise.allSettled(workers);

  return results;
}
