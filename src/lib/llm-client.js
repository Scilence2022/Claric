/**
 * Unified LLM Client Module
 *
 * Provides a shared abstraction for both Ollama and vLLM backends using the
 * OpenAI-compatible /v1/chat/completions format. All functions are pure --
 * they accept config objects and return promises with no global state.
 *
 * @module llm-client
 */

/**
 * Strips <think>...</think> tags and reasoning artifacts from LLM responses.
 * Applied to ALL backends as a universal safety net.
 *
 * Multi-pass regex strategy:
 *   Pass 1: Remove complete <think>...</think> blocks (multiline-safe)
 *   Pass 2: Remove orphaned </think> closing tags
 *   Pass 3: Remove orphaned <think> opening tags
 *   Pass 4: Trim whitespace and collapse excessive newlines
 *
 * @param {string} text - Raw LLM response text
 * @param {function} [log] - Optional logging callback (message, type)
 * @returns {string} Cleaned text with reasoning artifacts removed
 */
export function stripThinkTags(text, log) {
  if (!text) return text;

  let cleaned = text;
  let hadTags = false;

  // Pass 1: Strip complete <think>...</think> blocks (including multiline)
  const pass1 = cleaned.replace(/<think>[\s\S]*?<\/think>/gi, '');
  if (pass1 !== cleaned) hadTags = true;
  cleaned = pass1;

  // Pass 2: Strip orphaned </think> tags (closing without opening)
  const pass2 = cleaned.replace(/<\/think>/gi, '');
  if (pass2 !== cleaned) hadTags = true;
  cleaned = pass2;

  // Pass 3: Strip orphaned <think> tags (opening without closing)
  const pass3 = cleaned.replace(/<think>/gi, '');
  if (pass3 !== cleaned) hadTags = true;
  cleaned = pass3;

  // Pass 4: Trim whitespace and collapse 3+ newlines to 2
  cleaned = cleaned.trim().replace(/\n{3,}/g, '\n\n');

  if (hadTags && typeof log === 'function') {
    log('Cleaned reasoning artifacts from response', 'info');
  }

  return cleaned;
}

/**
 * Strips common markdown formatting artifacts from LLM responses.
 * Used as a post-processing safety net for amendment-mode responses
 * where the output will be inserted as plain text into Word documents.
 *
 * Removes:
 *   - Bold markers: **text** -> text
 *   - Italic markers: *text* -> text (but not bullet-point asterisks at line start)
 *   - Heading markers: ### heading -> heading
 *   - Inline code: `code` -> code
 *   - Code fences: ```...``` -> content only
 *
 * Does NOT remove:
 *   - Numbered lists (1. 2. 3.) -- these are common in contracts
 *   - Horizontal rules (---) -- could be intentional
 *   - Links [text](url) -- rare in contracts, leave as-is
 *
 * @param {string} text - LLM response text potentially containing markdown
 * @param {function} [log] - Optional logging callback (message, type)
 * @returns {string} Text with markdown formatting artifacts removed
 */
export function stripMarkdown(text, log) {
  if (!text) return text;

  let cleaned = text;
  let hadMarkdown = false;

  // Strip code fences (```language ... ```)
  const pass1 = cleaned.replace(/```[\s\S]*?```/g, (match) => {
    // Extract content between fences, removing the fence lines themselves
    const lines = match.split('\n');
    // Remove first line (```lang) and last line (```)
    return lines.slice(1, -1).join('\n');
  });
  if (pass1 !== cleaned) hadMarkdown = true;
  cleaned = pass1;

  // Strip heading markers (### at start of line)
  const pass2 = cleaned.replace(/^#{1,6}\s+/gm, '');
  if (pass2 !== cleaned) hadMarkdown = true;
  cleaned = pass2;

  // Strip bold markers (**text** or __text__)
  const pass3 = cleaned.replace(/\*\*(.+?)\*\*/g, '$1').replace(/__(.+?)__/g, '$1');
  if (pass3 !== cleaned) hadMarkdown = true;
  cleaned = pass3;

  // Strip italic markers (*text* or _text_) — but NOT bullet-point asterisks at line start
  // Only match *text* that is NOT at the start of a line (bullet points)
  const pass4 = cleaned.replace(/(?<!^)(?<![\n])\*([^\s*][^*]*?)\*/gm, '$1');
  if (pass4 !== cleaned) hadMarkdown = true;
  cleaned = pass4;

  // Strip bullet-point asterisks at line start (* item -> item)
  const pass5 = cleaned.replace(/^\*\s+/gm, '');
  if (pass5 !== cleaned) hadMarkdown = true;
  cleaned = pass5;

  // Strip inline code backticks (`code` -> code)
  const pass6 = cleaned.replace(/`([^`]+)`/g, '$1');
  if (pass6 !== cleaned) hadMarkdown = true;
  cleaned = pass6;

  if (hadMarkdown && typeof log === 'function') {
    log('Stripped markdown formatting from response', 'info');
  }

  return cleaned;
}

/**
 * Strips chunk delimiter markers that the LLM may echo from the prompt.
 *
 * During whole-document processing, the orchestrator wraps chunk text in
 * delimiter markers ([AMEND THIS TEXT]...[END TEXT] and
 * [CONTEXT - DO NOT AMEND]...[END CONTEXT]). LLMs sometimes echo these
 * markers in their output. This function removes them as a safety net.
 *
 * @param {string} text - LLM response text potentially containing delimiter markers
 * @param {function} [log] - Optional logging callback (message, type)
 * @returns {string} Text with delimiter markers removed
 */
export function stripChunkDelimiters(text, log) {
  if (!text) return text;

  let cleaned = text;
  let hadDelimiters = false;

  // Remove all four delimiter markers (case-insensitive, with optional surrounding whitespace on the line)
  const markers = [
    /^\s*\[AMEND THIS TEXT\]\s*$/gmi,
    /^\s*\[END TEXT\]\s*$/gmi,
    /^\s*\[CONTEXT\s*-\s*DO NOT AMEND\]\s*$/gmi,
    /^\s*\[END CONTEXT\]\s*$/gmi,
  ];

  for (const marker of markers) {
    const pass = cleaned.replace(marker, '');
    if (pass !== cleaned) hadDelimiters = true;
    cleaned = pass;
  }

  // Collapse any resulting triple+ newlines to double
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n').trim();

  if (hadDelimiters && typeof log === 'function') {
    log('Stripped chunk delimiter markers from response', 'info');
  }

  return cleaned;
}

/**
 * Incremental <think>...</think> demultiplexer for streamed tokens.
 *
 * Reasoning models emit their chain-of-thought either in a separate
 * `reasoning_content` field (handled by the SSE parsers) or inline in the
 * content wrapped in <think> tags. This demux routes streamed content tokens
 * to onContent and inline think-tag content to onReasoning as they arrive,
 * holding back partial tag suffixes across token boundaries.
 *
 * Mirrors stripThinkTags semantics: a never-closed <think> leaves its text in
 * the reasoning channel (the tag itself is never emitted to content).
 *
 * @param {object} handlers
 * @param {function} [handlers.onContent] - Receives answer-text deltas
 * @param {function} [handlers.onReasoning] - Receives thinking deltas
 * @returns {{ push: function(string), flush: function() }}
 */
export function createStreamDemux({ onContent, onReasoning } = {}) {
  let buffer = '';
  let inThink = false;

  const emit = (text, toReasoning) => {
    if (!text) return;
    if (toReasoning) {
      if (typeof onReasoning === 'function') onReasoning(text);
    } else if (typeof onContent === 'function') {
      onContent(text);
    }
  };

  // Longest k such that `text` ends with a k-char prefix of `tag`
  // (case-insensitive) — those chars may be the start of a tag split
  // across tokens and must be held back for the next push.
  function suffixHold(text, tag) {
    const lower = text.toLowerCase();
    for (let k = Math.min(tag.length - 1, text.length); k > 0; k--) {
      if (lower.endsWith(tag.slice(0, k))) return k;
    }
    return 0;
  }

  return {
    push(token) {
      buffer += token;
      for (;;) {
        const tag = inThink ? '</think>' : '<think>';
        const idx = buffer.toLowerCase().indexOf(tag);
        if (idx === -1) {
          const hold = suffixHold(buffer, tag);
          emit(buffer.slice(0, buffer.length - hold), inThink);
          buffer = buffer.slice(buffer.length - hold);
          return;
        }
        emit(buffer.slice(0, idx), inThink);
        buffer = buffer.slice(idx + tag.length);
        inThink = !inThink;
      }
    },
    flush() {
      if (buffer) emit(buffer, inThink);
      buffer = '';
      inThink = false;
    },
  };
}

/**
 * Joins a base URL, the provider's API prefix, and an endpoint path.
 * Trailing slashes on the base are stripped, and a base that ALREADY ends
 * with the API prefix is not suffixed again — users entering the full
 * endpoint (e.g. https://host/v1 with the default apiPath) get
 * https://host/v1/chat/completions, not a double /v1/v1 404.
 *
 * @param {string} baseUrl - Configured endpoint (proxy path or absolute URL)
 * @param {string} apiPath - API prefix (e.g. '/v1' or '/api/paas/v4')
 * @param {string} endpoint - Final path segment (e.g. '/chat/completions')
 * @returns {string}
 * @private
 */
function joinApiUrl(baseUrl, apiPath, endpoint) {
  const base = String(baseUrl || '').replace(/\/+$/, '');
  const prefix = apiPath.startsWith('/') ? apiPath : `/${apiPath}`;
  const withPrefix = base.endsWith(prefix) ? base : base + prefix;
  return withPrefix + endpoint;
}

/**
 * Private helper to build the request URL and headers for chat completions.
 *
 * The API prefix comes from config.apiPath (default '/v1'). Most providers
 * serve OpenAI-compatible endpoints under /v1, but some use a different
 * prefix (e.g. Zhipu GLM: /api/paas/v4) -- configured per provider preset.
 *
 * @param {object} config - Backend configuration
 * @returns {{ url: string, headers: object }}
 */
function buildRequestConfig(config) {
  const url = joinApiUrl(config.url, config.apiPath || '/v1', '/chat/completions');
  const headers = { 'Content-Type': 'application/json' };
  if (config.apiKey) {
    headers['Authorization'] = `Bearer ${config.apiKey}`;
  }
  return { url, headers };
}

/**
 * Throws when a finished response was cut off by the model's token limit.
 * A length-truncated amendment or comment flowing into the diff pipeline
 * would be applied as if it were complete — refuse instead.
 *
 * @param {{finish_reason?: string}} [choice]
 * @private
 */
function _assertNotLengthTruncated(choice) {
  if (choice && choice.finish_reason === 'length') {
    throw new Error(
      'LLM output truncated (finish_reason=length): the response hit the model\'s ' +
      'max token limit. Reduce the scope or selection and retry.'
    );
  }
}

/**
 * Builds the error message for a non-ok response, including a truncated
 * slice of the response body: OpenAI-compatible backends put the actionable
 * detail there (rate-limit reason, unknown model, context-length exceeded),
 * and HTTP/2 often leaves statusText empty — without the body the user just
 * sees "HTTP 429: ". The body is server-controlled text shown to the user,
 * not logged server-side, and never contains the request's credentials.
 *
 * @param {Response} response - A response with ok === false
 * @returns {Promise<string>} e.g. "HTTP 429: {"error":{"message":"rate limited"}}"
 * @private
 */
async function _describeHttpError(response) {
  let detail = '';
  try {
    detail = (await response.text()).slice(0, 300).trim();
  } catch {
    // Unreadable body -- the status line alone still beats nothing.
  }
  const statusLine = `HTTP ${response.status}${response.statusText ? ` ${response.statusText}` : ''}`;
  return detail ? `${statusLine}: ${detail}` : statusLine;
}

/**
 * Sends a prompt to the configured LLM backend.
 * Uses OpenAI-compatible /v1/chat/completions format for both Ollama and vLLM.
 *
 * @param {object} config - Backend configuration
 * @param {string} config.url - Base proxy path (e.g., '/ollama' or '/vllm')
 * @param {string} config.apiKey - API key (empty string if not required)
 * @param {string} config.model - Model identifier
 * @param {string} promptText - The prompt text to send
 * @param {function} [log] - Optional logging callback (message, type)
 * @returns {Promise<string>} The LLM response text with think tags stripped
 * @throws {Error} On non-ok HTTP response or network failure
 */
/**
 * Sends a single-string prompt to the LLM backend as a one-message
 * user-role chat completion. Kept for callers that need the legacy
 * shim shape; new code should prefer {@link sendMessages}.
 *
 * Abort/timeout wiring mirrors sendMessages (WebView2-safe, no
 * AbortSignal.any): the optional external signal aborts the local
 * controller, which aborts the fetch. Aborts from timeout and from the
 * external signal are reported with distinct error names.
 *
 * @param {Object} config - { url, apiKey, model }
 * @param {string} promptText - User-role prompt body
 * @param {function} [log] - Optional logging callback (message, type)
 * @param {AbortSignal} [signal] - Optional abort signal for cancellation
 * @param {number} [timeoutMs=120000] - Per-request timeout in ms
 * @returns {Promise<string>} Cleaned LLM response text
 * @throws {DOMException} AbortError on user cancellation via signal
 * @throws {Error} TimeoutError (error.name === 'TimeoutError') when timeout expires
 */
export async function sendPrompt(config, promptText, log, signal, timeoutMs = 120000) {
  const { url, headers } = buildRequestConfig(config);

  const body = JSON.stringify({
    model: config.model,
    messages: [{ role: 'user', content: promptText }],
    stream: false,
  });

  const localController = new AbortController();
  let timedOut = false;
  const timeoutId = setTimeout(() => {
    timedOut = true;
    localController.abort();
  }, timeoutMs);

  let onExternalAbort;
  if (signal) {
    if (signal.aborted) {
      clearTimeout(timeoutId);
      throw new DOMException('The operation was aborted.', 'AbortError');
    }
    onExternalAbort = () => localController.abort();
    signal.addEventListener('abort', onExternalAbort);
  }

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body,
      signal: localController.signal,
    });

    if (!response.ok) {
      throw new Error(await _describeHttpError(response));
    }

    const data = await response.json();
    const choice = data.choices?.[0];
    _assertNotLengthTruncated(choice);
    const rawText = choice?.message?.content ?? '';
    return stripThinkTags(rawText, log);
  } catch (err) {
    if (timedOut && err.name === 'AbortError') {
      const timeoutErr = new Error(`LLM request timed out after ${Math.round(timeoutMs / 1000)}s`);
      timeoutErr.name = 'TimeoutError';
      throw timeoutErr;
    }
    throw err;
  } finally {
    clearTimeout(timeoutId);
    if (signal && onExternalAbort) {
      signal.removeEventListener('abort', onExternalAbort);
    }
  }
}

/**
 * Sends a messages array to the LLM backend, preserving system/user roles.
 * Unlike sendPrompt (single string), this sends the messages array directly
 * to the chat completions API without wrapping in a single user message.
 *
 * Uses a manual AbortController approach instead of AbortSignal.any() for
 * compatibility with Office's WebView2 runtime.
 *
 * @param {Object} config - { url, apiKey, model }
 * @param {Array<{role: string, content: string}>} messages - Chat messages
 * @param {function} [log] - Optional logging callback (message, type)
 * @param {AbortSignal} [signal] - Optional abort signal for cancellation
 * @param {number} [timeoutMs=120000] - Per-request timeout in ms
 * @returns {Promise<string>} Cleaned LLM response text
 * @throws {Error} On non-ok HTTP response or network failure
 * @throws {DOMException} AbortError on user cancellation via signal
 * @throws {Error} TimeoutError (error.name === 'TimeoutError') when timeout expires
 */
export async function sendMessages(config, messages, log, signal, timeoutMs = 120000) {
  const { url, headers } = buildRequestConfig(config);

  const body = JSON.stringify({
    model: config.model,
    messages: messages,
    stream: false,
  });

  // Create a local AbortController for timeout management
  const localController = new AbortController();
  let timedOut = false;
  const timeoutId = setTimeout(() => {
    timedOut = true;
    localController.abort();
  }, timeoutMs);

  // Wire external signal to trigger local abort (WebView2-safe, no AbortSignal.any)
  let onExternalAbort;
  if (signal) {
    if (signal.aborted) {
      clearTimeout(timeoutId);
      throw new DOMException('The operation was aborted.', 'AbortError');
    }
    onExternalAbort = () => localController.abort();
    signal.addEventListener('abort', onExternalAbort);
  }

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body,
      signal: localController.signal,
    });

    if (!response.ok) {
      throw new Error(await _describeHttpError(response));
    }

    const data = await response.json();
    const choice = data.choices?.[0];
    _assertNotLengthTruncated(choice);
    const rawText = choice?.message?.content ?? '';
    return stripThinkTags(rawText, log);
  } catch (err) {
    // Distinguish timeout aborts from user cancellation aborts
    if (timedOut && err.name === 'AbortError') {
      const timeoutErr = new Error(`LLM request timed out after ${Math.round(timeoutMs / 1000)}s`);
      timeoutErr.name = 'TimeoutError';
      throw timeoutErr;
    }
    throw err;
  } finally {
    clearTimeout(timeoutId);
    if (signal && onExternalAbort) {
      signal.removeEventListener('abort', onExternalAbort);
    }
  }
}

/**
 * Sends a messages array to the LLM backend with OpenAI-compatible streaming
 * (`stream: true`), preserving system/user roles. Content deltas are demuxed
 * through createStreamDemux: reasoning (the `reasoning_content` field plus
 * inline <think> blocks) streams to handlers.onReasoning, answer text to
 * handlers.onContent.
 *
 * If the server ignores `stream: true` and answers with a plain JSON body
 * (non-SSE content type), falls back to reading the whole response and
 * delivering it as a single delta -- so callers get correct output from
 * backends without streaming support.
 *
 * Abort/timeout wiring mirrors sendMessages (WebView2-safe, no AbortSignal.any).
 *
 * @param {Object} config - { url, apiKey, model, apiPath }
 * @param {Array<{role: string, content: string}>} messages - Chat messages
 * @param {function|{onContent?: function, onReasoning?: function}} [handlers] -
 *   A plain function is treated as onContent (legacy shorthand)
 * @param {function} [log] - Optional logging callback (message, type)
 * @param {AbortSignal} [signal] - Optional abort signal for cancellation
 * @param {number} [timeoutMs=120000] - Idle timeout in ms: the request aborts
 *   only when the backend sends no data for this long (a long but actively
 *   streaming generation never trips it)
 * @returns {Promise<{content: string, reasoning: string}>} Full text per channel
 *   (content is think-tag stripped)
 * @throws {Error} On non-ok HTTP response or network failure
 * @throws {DOMException} AbortError on user cancellation via signal
 * @throws {Error} TimeoutError (error.name === 'TimeoutError') when timeout expires
 */
export async function sendMessagesStream(config, messages, handlers, log, signal, timeoutMs = 120000) {
  const { url, headers } = buildRequestConfig(config);
  const onContent = typeof handlers === 'function' ? handlers : handlers?.onContent;
  const onReasoning = typeof handlers === 'function' ? undefined : handlers?.onReasoning;
  let full = '';
  let reasoning = '';
  const demux = createStreamDemux({
    onContent: (t) => { full += t; if (onContent) onContent(t); },
    onReasoning: (t) => { reasoning += t; if (onReasoning) onReasoning(t); },
  });

  const body = JSON.stringify({
    model: config.model,
    messages: messages,
    stream: true,
  });

  const localController = new AbortController();
  let timedOut = false;
  // Idle timeout: the clock resets whenever the backend sends data, so a
  // long-but-active generation (e.g. a large SVG illustration) never trips
  // it -- only a genuinely stalled stream does.
  let timeoutId;
  const armIdleTimeout = () => {
    clearTimeout(timeoutId);
    timeoutId = setTimeout(() => {
      timedOut = true;
      localController.abort();
    }, timeoutMs);
  };
  armIdleTimeout();

  let onExternalAbort;
  if (signal) {
    if (signal.aborted) {
      clearTimeout(timeoutId);
      throw new DOMException('The operation was aborted.', 'AbortError');
    }
    onExternalAbort = () => localController.abort();
    signal.addEventListener('abort', onExternalAbort);
  }

  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body,
      signal: localController.signal,
    });

    if (!response.ok) {
      throw new Error(await _describeHttpError(response));
    }

    // Headers arrived: the backend is alive -- restart the idle clock.
    armIdleTimeout();

    const contentType = response.headers && typeof response.headers.get === 'function'
      ? (response.headers.get('content-type') || '')
      : '';

    // Non-SSE fallback: the backend ignored stream:true and sent plain JSON.
    if (!contentType.includes('text/event-stream') || !response.body || typeof response.body.getReader !== 'function') {
      const data = await response.json();
      _assertNotLengthTruncated(data.choices?.[0]);
      const message = data.choices?.[0]?.message ?? {};
      const reasoningText = message.reasoning_content ?? '';
      if (reasoningText) {
        reasoning += reasoningText;
        if (onReasoning) onReasoning(reasoningText);
      }
      demux.push(message.content ?? '');
      demux.flush();
      return { content: stripThinkTags(full, log), reasoning: reasoning.trim() };
    }

    // SSE parsing: buffer partial lines across chunks, handle `data:` lines
    // and the `[DONE]` terminator.
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let doneReceived = false;
    let finishReason = null;

    const handleDataLine = (line) => {
      if (doneReceived) return; // stop honoring deltas after [DONE]
      const payload = line.slice(5).trim();
      if (payload === '[DONE]') {
        doneReceived = true;
        return;
      }
      if (!payload) return;
      try {
        const json = JSON.parse(payload);
        const choice = json.choices?.[0];
        const delta = choice?.delta ?? choice?.message ?? {};
        if (choice?.finish_reason) finishReason = choice.finish_reason;
        const reasoningToken = delta.reasoning_content ?? '';
        if (reasoningToken) {
          reasoning += reasoningToken;
          if (onReasoning) onReasoning(reasoningToken);
        }
        const token = delta.content ?? '';
        if (token) demux.push(token);
      } catch (_parseErr) {
        // Incomplete or non-JSON data line -- skip it.
      }
    };

    const drainBuffer = () => {
      let newlineIdx;
      while (!doneReceived && (newlineIdx = buffer.indexOf('\n')) !== -1) {
        const line = buffer.slice(0, newlineIdx).trim();
        buffer = buffer.slice(newlineIdx + 1);
        if (!line.startsWith('data:')) continue;
        handleDataLine(line);
      }
    };

    while (!doneReceived) {
      const { done, value } = await reader.read();
      if (done) break;
      // Data is flowing: restart the idle clock on every received chunk.
      armIdleTimeout();
      buffer += decoder.decode(value, { stream: true });
      drainBuffer();
    }
    // Flush the decoder's multi-byte tail, then drain a final data line
    // that may arrive without a trailing newline.
    buffer += decoder.decode();
    drainBuffer();
    if (!doneReceived && buffer.startsWith('data:')) {
      handleDataLine(buffer.trim());
      buffer = '';
    }
    if (doneReceived && typeof reader.cancel === 'function') {
      // [DONE] received: release the keep-alive body instead of leaving
      // the reader open.
      Promise.resolve(reader.cancel()).catch(() => {});
    }
    if (!doneReceived && !finishReason) {
      // Neither terminator arrived: the stream was cut short (proxy close,
      // network drop). Returning the partial text would present a half
      // answer — or worse, apply half an amendment — as if complete.
      throw new Error(
        'LLM stream closed before completion (no [DONE] marker or finish_reason) — ' +
        'the output may be truncated. Retry the request.'
      );
    }
    // A stream that terminates cleanly but with finish_reason=length is
    // still truncated: the model hit its max token limit mid-amendment.
    // Same refusal as the non-streaming paths, which the SSE path skipped.
    _assertNotLengthTruncated({ finish_reason: finishReason });

    demux.flush();
    return { content: stripThinkTags(full, log), reasoning: reasoning.trim() };
  } catch (err) {
    if (timedOut && err.name === 'AbortError') {
      const timeoutErr = new Error(`LLM request timed out: no output from the model for ${Math.round(timeoutMs / 1000)}s`);
      timeoutErr.name = 'TimeoutError';
      throw timeoutErr;
    }
    throw err;
  } finally {
    clearTimeout(timeoutId);
    if (signal && onExternalAbort) {
      signal.removeEventListener('abort', onExternalAbort);
    }
  }
}

/**
 * Sends a prompt to the LLM backend with streaming. Thin wrapper over
 * sendMessagesStream for single-string prompts.
 *
 * @param {Object} config - { url, apiKey, model, apiPath }
 * @param {string} promptText - The prompt text to send
 * @param {function|{onContent?: function, onReasoning?: function}} [handlers] -
 *   A plain function is treated as onContent (legacy shorthand)
 * @param {function} [log] - Optional logging callback (message, type)
 * @param {AbortSignal} [signal] - Optional abort signal for cancellation
 * @param {number} [timeoutMs=120000] - Per-request timeout in ms
 * @returns {Promise<string>} The full cleaned response text (think tags stripped)
 */
export async function sendPromptStream(config, promptText, handlers, log, signal, timeoutMs = 120000) {
  const { content } = await sendMessagesStream(
    config,
    [{ role: 'user', content: promptText }],
    handlers,
    log,
    signal,
    timeoutMs
  );
  return content;
}

// Bounds the Settings-page connection probe: a hung backend must surface a
// reportable timeout instead of leaving the status dot on "Connecting"
// forever. Generous enough for a cold-starting local Ollama.
const CONNECTION_TEST_TIMEOUT_MS = 30000;

/**
 * Tests connection to the configured LLM backend and retrieves model list.
 * Uses the OpenAI-compatible models endpoint (prefix from config.apiPath,
 * default /v1) for all providers.
 *
 * @param {object} config - Backend configuration
 * @param {string} config.url - Base URL (proxy path or provider origin)
 * @param {string} [config.apiPath='/v1'] - API prefix (e.g. '/api/paas/v4')
 * @param {string} config.apiKey - API key (empty string if not required)
 * @returns {Promise<{connected: boolean, models: Array<{id: string}>}>}
 * @throws {Error} On non-ok HTTP response or network failure
 * @throws {Error} TimeoutError (error.name === 'TimeoutError') when the
 *   backend does not answer within 30s
 */
export async function testConnection(config) {
  const url = joinApiUrl(config.url, config.apiPath || '/v1', '/models');

  const headers = { Accept: 'application/json' };
  if (config.apiKey) {
    headers['Authorization'] = `Bearer ${config.apiKey}`;
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), CONNECTION_TEST_TIMEOUT_MS);

  let response;
  try {
    response = await fetch(url, { method: 'GET', headers, signal: controller.signal });
  } catch (err) {
    if (err.name === 'AbortError') {
      const timeoutErr = new Error(
        `Connection test timed out after ${CONNECTION_TEST_TIMEOUT_MS / 1000}s`
      );
      timeoutErr.name = 'TimeoutError';
      throw timeoutErr;
    }
    throw err;
  } finally {
    clearTimeout(timeoutId);
  }

  if (!response.ok) {
    throw new Error(await _describeHttpError(response));
  }

  const data = await response.json();
  const models = (data.data || []).map((m) => ({ id: m.id }));
  return { connected: true, models };
}

