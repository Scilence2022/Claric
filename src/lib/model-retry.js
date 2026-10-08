/** Bounded retries for model generation only; document writes never enter this loop. */
const TRANSIENT_HTTP_STATUSES = new Set([408, 429, 500, 502, 503, 504, 529]);
const DEFAULT_MAX_RETRIES = 2;
const MAX_RETRY_WAIT_MS = 60000;

function abortError() {
  return new DOMException('The operation was aborted.', 'AbortError');
}

/** Wait without blocking cancellation, and remove the listener on every exit. */
function waitForRetry(delayMs, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) { reject(abortError()); return; }
    const onAbort = () => {
      clearTimeout(timer);
      signal.removeEventListener('abort', onAbort);
      reject(abortError());
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener('abort', onAbort);
      resolve(undefined);
    }, delayMs);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

/**
 * Retry transient generation failures with exponential backoff and jitter.
 * Each operation invocation owns a fresh request timeout/controller. A stream
 * caller can refuse retries once any answer or reasoning was delivered.
 *
 * @param {function(): Promise<any>} operation
 * @param {object} [options]
 * @param {number} [options.maxRetries=2] - Additional attempts, bounded to 0..5
 * @param {AbortSignal} [options.signal]
 * @param {function} [options.log]
 * @param {function(): boolean} [options.canRetry]
 * @returns {Promise<any>}
 */
export async function withModelRetry(operation, { maxRetries, signal, log, canRetry = () => true } = {}) {
  const retries = Number.isInteger(maxRetries) && maxRetries >= 0
    ? Math.min(maxRetries, 5) : DEFAULT_MAX_RETRIES;
  for (let attempt = 0; ; attempt++) {
    if (signal?.aborted) throw abortError();
    try {
      const result = await operation();
      if (signal?.aborted) throw abortError();
      return result;
    } catch (error) {
      // Cancellation always wins, including a race with an attempt's timeout.
      if (signal?.aborted) throw abortError();
      const transient = error.name !== 'AbortError' && (
        TRANSIENT_HTTP_STATUSES.has(error.status) || error.name === 'TimeoutError' || error.retryable === true
      );
      if (!transient || attempt >= retries || !canRetry() || error.retryAfterMs > MAX_RETRY_WAIT_MS) throw error;
      const backoffMs = Math.min(30000, 1000 * (2 ** attempt) * (1 + Math.random() * 0.25));
      const delayMs = Math.max(backoffMs, error.retryAfterMs || 0);
      // Do not repeat provider-controlled response bodies or credentials in retry logs.
      const reason = error.status ? `HTTP ${error.status}` : error.name === 'TimeoutError' ? 'timeout' : 'connection interrupted';
      if (typeof log === 'function') {
        log(`LLM request failed (${reason}); retrying (${attempt + 1}/${retries}) in ${(delayMs / 1000).toFixed(1)}s.`, 'warning');
      }
      await waitForRetry(delayMs, signal);
    }
  }
}
