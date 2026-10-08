const { sendPrompt, sendMessages, sendPromptStream, sendMessagesStream } = require('../src/lib/llm-client.js');
const { withModelRetry } = require('../src/lib/model-retry.js');
const { processChunksParallel } = require('../src/lib/orchestrator.js');

const CONFIG = { url: 'https://model.example', model: 'test', apiKey: 'secret', temperature: 0.4 };
const CLAUDE = { ...CONFIG, provider: 'claude', model: 'claude-sonnet-4-6' };
const MESSAGES = [{ role: 'system', content: 'Preserve meaning.' }, { role: 'user', content: 'Polish this text.' }];
const observe = (promise) => promise.then(value => ({ value }), error => ({ error }));

function httpError(status = 504, retryAfter) {
  return {
    ok: false, status, statusText: '',
    headers: { get: () => retryAfter },
    text: async () => '{"error":{"message":"The upstream service timed out.","code":"upstream_timeout"}}',
  };
}

function success(config = CONFIG) {
  return {
    ok: true, headers: { get: () => 'application/json' },
    json: async () => config.provider === 'claude'
      ? { content: [{ type: 'text', text: 'Polished text.' }], stop_reason: 'end_turn' }
      : { choices: [{ message: { content: 'Polished text.' }, finish_reason: 'stop' }] },
  };
}

function sse(chunks, failure) {
  let index = 0;
  return {
    ok: true, headers: { get: () => 'text/event-stream' },
    body: { getReader: () => ({ read: async () => {
      if (index < chunks.length) return { done: false, value: new TextEncoder().encode(chunks[index++]) };
      if (failure) throw failure;
      return { done: true };
    } }) },
  };
}
const contentLine = (text) => `data: ${JSON.stringify({ choices: [{ delta: { content: text } }] })}\n`;
const reasoningLine = (text) => `data: ${JSON.stringify({ choices: [{ delta: { reasoning_content: text } }] })}\n`;

beforeEach(() => {
  jest.useFakeTimers();
  jest.setSystemTime(new Date('2026-10-08T00:00:00Z'));
  // Midpoint jitter also keeps source-map's randomized quicksort balanced.
  jest.spyOn(Math, 'random').mockReturnValue(0.5);
  global.fetch = jest.fn();
});
afterEach(() => {
  jest.useRealTimers();
  jest.restoreAllMocks();
  delete global.fetch;
});

describe('generation retry policy', () => {
  test.each([
    ['prompt', sendPrompt, 'Polish this text.'],
    ['messages', sendMessages, MESSAGES],
    ['stream prompt', sendPromptStream, 'Polish this text.'],
    ['stream messages', sendMessagesStream, MESSAGES],
  ])('%s retries HTTP 504 with identical inputs on both providers', async (_name, send, input) => {
    for (const config of [CONFIG, CLAUDE]) {
      global.fetch.mockReset().mockResolvedValueOnce(httpError()).mockResolvedValue(success(config));
      const outcome = observe(send(config, input));
      await jest.runAllTimersAsync();
      const { value, error } = await outcome;
      expect(error).toBeUndefined();
      expect(typeof value === 'string' ? value : value.content).toBe('Polished text.');
      expect(global.fetch).toHaveBeenCalledTimes(2);
      const [first, second] = global.fetch.mock.calls;
      expect(second[0]).toBe(first[0]);
      expect(second[1].body).toBe(first[1].body);
      expect(second[1].headers).toEqual(first[1].headers);
      expect(second[1].signal).not.toBe(first[1].signal);
      expect(jest.getTimerCount()).toBe(0);
    }
  });

  test.each([408, 429, 500, 502, 503, 504, 529])('recovers from transient HTTP %s', async (status) => {
    global.fetch.mockResolvedValueOnce(httpError(status)).mockResolvedValue(success());
    const outcome = observe(sendMessages(CONFIG, MESSAGES));
    await jest.runAllTimersAsync();
    expect(await outcome).toEqual({ value: 'Polished text.' });
    expect(global.fetch).toHaveBeenCalledTimes(2);
  });

  test('exhausts exactly two retries, logs delays, and preserves the final provider error', async () => {
    global.fetch.mockResolvedValue(httpError());
    const log = jest.fn();
    const outcome = observe(sendMessages(CONFIG, MESSAGES, log));
    await jest.runAllTimersAsync();
    const { error } = await outcome;
    expect(error.message).toContain('HTTP 504:');
    expect(error.message).toContain('upstream_timeout');
    expect(global.fetch).toHaveBeenCalledTimes(3);
    expect(log.mock.calls).toEqual([
      ['LLM request failed (HTTP 504); retrying (1/2) in 1.1s.', 'warning'],
      ['LLM request failed (HTTP 504); retrying (2/2) in 2.3s.', 'warning'],
    ]);
    expect(JSON.stringify(log.mock.calls)).not.toContain('secret');
    expect(jest.getTimerCount()).toBe(0);
  });

  test.each([400, 401, 403, 404, 413, 422, 501])('does not retry permanent HTTP %s', async (status) => {
    global.fetch.mockResolvedValue(httpError(status));
    await expect(sendMessages(CONFIG, MESSAGES)).rejects.toMatchObject({ status });
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test.each([TypeError, class NetworkError extends Error { constructor() { super('Offline'); this.name = 'NetworkError'; } }])(
    'recovers from a browser network failure (%s)', async (ErrorType) => {
      global.fetch.mockRejectedValueOnce(new ErrorType('Failed to fetch')).mockResolvedValue(success());
      const outcome = observe(sendMessages(CONFIG, MESSAGES));
      await jest.runAllTimersAsync();
      expect(await outcome).toEqual({ value: 'Polished text.' });
    }
  );

  test('recovers from a network error while reading the response body', async () => {
    global.fetch.mockResolvedValueOnce({ ok: true, json: async () => { throw new TypeError('terminated'); } })
      .mockResolvedValue(success());
    const outcome = observe(sendMessages(CONFIG, MESSAGES));
    await jest.runAllTimersAsync();
    expect(await outcome).toEqual({ value: 'Polished text.' });
  });

  test.each([CONFIG, CLAUDE])('retries local timeouts with a fresh per-attempt timeout (%s)', async (config) => {
    for (const send of [sendMessages, sendMessagesStream]) {
      global.fetch.mockReset().mockImplementationOnce((_url, { signal }) => new Promise((_resolve, reject) => {
        signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')));
      })).mockResolvedValue(success(config));
      const outcome = observe(send === sendMessages
        ? send(config, MESSAGES, undefined, undefined, 100)
        : send(config, MESSAGES, undefined, undefined, undefined, 100));
      await jest.runAllTimersAsync();
      expect(await outcome).toEqual({ value: send === sendMessages ? 'Polished text.' : { content: 'Polished text.', reasoning: '' } });
      expect(global.fetch).toHaveBeenCalledTimes(2);
      expect(global.fetch.mock.calls[0][1].signal.aborted).toBe(true);
      expect(global.fetch.mock.calls[1][1].signal.aborted).toBe(false);
      expect(jest.getTimerCount()).toBe(0);
    }
  });

  test.each(['seconds', 'date'])('respects Retry-After %s', async (format) => {
    const retryAfter = format === 'seconds' ? '3' : new Date(Date.now() + 3000).toUTCString();
    global.fetch.mockResolvedValueOnce(httpError(429, retryAfter)).mockResolvedValue(success());
    const outcome = observe(sendMessages(CONFIG, MESSAGES));
    await jest.advanceTimersByTimeAsync(2999);
    expect(global.fetch).toHaveBeenCalledTimes(1);
    await jest.advanceTimersByTimeAsync(1);
    expect(await outcome).toEqual({ value: 'Polished text.' });
  });

  test.each(['invalid', '-2', '0'])('uses backoff for invalid/past Retry-After %s', async (header) => {
    global.fetch.mockResolvedValueOnce(httpError(503, header)).mockResolvedValue(success());
    const outcome = observe(sendMessages(CONFIG, MESSAGES));
    await jest.advanceTimersByTimeAsync(1124);
    expect(global.fetch).toHaveBeenCalledTimes(1);
    await jest.advanceTimersByTimeAsync(1);
    expect(await outcome).toEqual({ value: 'Polished text.' });
  });

  test('leaves long provider cooldowns for manual retry', async () => {
    global.fetch.mockResolvedValue(httpError(429, '120'));
    await expect(sendMessages(CONFIG, MESSAGES)).rejects.toMatchObject({ status: 429, retryAfterMs: 120000 });
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('does not retry token-limit truncation or malformed JSON', async () => {
    global.fetch.mockResolvedValue({ ok: true, json: async () => ({ choices: [{ finish_reason: 'length' }] }) });
    await expect(sendMessages(CONFIG, MESSAGES)).rejects.toThrow('truncated');
    expect(global.fetch).toHaveBeenCalledTimes(1);
    global.fetch.mockReset().mockResolvedValue({ ok: true, json: async () => { throw new SyntaxError('Invalid JSON'); } });
    await expect(sendMessages(CONFIG, MESSAGES)).rejects.toThrow('Invalid JSON');
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });
});

describe('cancellation and retry budgets', () => {
  test('cancel during a retry wait prevents the next request and cleans listeners', async () => {
    const controller = new AbortController();
    const add = jest.spyOn(controller.signal, 'addEventListener');
    const remove = jest.spyOn(controller.signal, 'removeEventListener');
    global.fetch.mockResolvedValue(httpError());
    const outcome = observe(sendMessages(CONFIG, MESSAGES, undefined, controller.signal));
    await jest.advanceTimersByTimeAsync(0);
    controller.abort();
    expect((await outcome).error.name).toBe('AbortError');
    await jest.runAllTimersAsync();
    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(remove.mock.calls.map(call => call[1])).toEqual(add.mock.calls.map(call => call[1]));
    expect(jest.getTimerCount()).toBe(0);
  });

  test('an active cancellation never retries', async () => {
    const controller = new AbortController();
    global.fetch.mockImplementation((_url, { signal }) => new Promise((_resolve, reject) => {
      signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')));
    }));
    const outcome = observe(sendMessages(CONFIG, MESSAGES, undefined, controller.signal, 100));
    await jest.advanceTimersByTimeAsync(99);
    controller.abort();
    expect((await outcome).error.name).toBe('AbortError');
    await jest.runAllTimersAsync();
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('cancellation wins when the local timeout fires in the same turn', async () => {
    const controller = new AbortController();
    global.fetch.mockImplementation((_url, { signal }) => new Promise((_resolve, reject) => {
      signal.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')));
    }));
    const outcome = observe(sendMessages(CONFIG, MESSAGES, undefined, controller.signal, 100));
    jest.advanceTimersByTime(100);
    controller.abort();
    expect((await outcome).error.name).toBe('AbortError');
    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(jest.getTimerCount()).toBe(0);
  });

  test('cancellation from the retry notification does not arm a wait timer', async () => {
    const controller = new AbortController();
    global.fetch.mockResolvedValue(httpError());
    await expect(sendMessages(CONFIG, MESSAGES, () => controller.abort(), controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(jest.getTimerCount()).toBe(0);
  });

  test('does not return a result that settles after cancellation', async () => {
    const controller = new AbortController();
    await expect(withModelRetry(async () => { controller.abort(); return 'Late result'; }, { signal: controller.signal }))
      .rejects.toMatchObject({ name: 'AbortError' });
  });

  test.each([0, 1, 10, -1, 1.5])('maxRetries %s is disabled, overridden, bounded, or defaulted', async (maxRetries) => {
    const operation = jest.fn().mockRejectedValue(Object.assign(new Error('Temporary'), { status: 504 }));
    const outcome = observe(withModelRetry(operation, { maxRetries }));
    await jest.runAllTimersAsync();
    expect((await outcome).error.message).toBe('Temporary');
    expect(operation).toHaveBeenCalledTimes(maxRetries === 0 ? 1 : maxRetries === 1 ? 2 : maxRetries === 10 ? 6 : 3);
  });

  test('jitter spreads retries and arbitrary application TypeErrors never retry', async () => {
    jest.mocked(Math.random).mockReturnValue(0.8);
    global.fetch.mockResolvedValueOnce(httpError()).mockResolvedValue(success());
    const outcome = observe(sendMessages(CONFIG, MESSAGES));
    await jest.advanceTimersByTimeAsync(1199);
    expect(global.fetch).toHaveBeenCalledTimes(1);
    await jest.advanceTimersByTimeAsync(1);
    expect(await outcome).toEqual({ value: 'Polished text.' });
    const operation = jest.fn().mockRejectedValue(new TypeError('Application bug'));
    await expect(withModelRetry(operation)).rejects.toThrow('Application bug');
    expect(operation).toHaveBeenCalledTimes(1);
  });
});

describe('stream replay boundaries', () => {
  test.each(['content', 'reasoning'])('does not replay after %s tokens reach the caller', async (kind) => {
    const tokens = jest.fn();
    const line = kind === 'content' ? contentLine : reasoningLine;
    global.fetch.mockResolvedValue(sse([line('Partial')], new TypeError('terminated')));
    await expect(sendMessagesStream(CONFIG, MESSAGES, { onContent: tokens, onReasoning: tokens }))
      .rejects.toThrow('terminated');
    expect(tokens).toHaveBeenCalledTimes(1);
    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(global.fetch.mock.calls[0][1].signal.aborted).toBe(true);
  });

  test('retries a stream that closes before output, and emits only the successful attempt', async () => {
    global.fetch.mockResolvedValueOnce(sse([])).mockResolvedValue(sse([contentLine('Complete'), 'data: [DONE]\n']));
    const tokens = jest.fn();
    const outcome = observe(sendMessagesStream(CONFIG, MESSAGES, tokens));
    await jest.runAllTimersAsync();
    expect(await outcome).toEqual({ value: { content: 'Complete', reasoning: '' } });
    expect(tokens.mock.calls).toEqual([['Complete']]);
    expect(global.fetch).toHaveBeenCalledTimes(2);
  });

  test.each([CONFIG, CLAUDE])('retries transient SSE error events before output (%s)', async (config) => {
    const payload = { type: 'error', error: { type: 'overloaded_error', message: 'Busy' } };
    global.fetch.mockResolvedValueOnce(sse([`data: ${JSON.stringify(payload)}\n`])).mockResolvedValue(success(config));
    const outcome = observe(sendMessagesStream(config, MESSAGES));
    await jest.runAllTimersAsync();
    expect(await outcome).toEqual({ value: { content: 'Polished text.', reasoning: '' } });
    expect(global.fetch).toHaveBeenCalledTimes(2);
  });

  test('does not replay Anthropic thinking after a later error event', async () => {
    const tokens = jest.fn();
    global.fetch.mockResolvedValue(sse([
      'data: {"type":"content_block_delta","delta":{"type":"thinking_delta","thinking":"Partial reasoning"}}\n',
      'data: {"type":"error","error":{"type":"overloaded_error","message":"Busy"}}\n',
    ]));
    await expect(sendMessagesStream(CLAUDE, MESSAGES, { onReasoning: tokens })).rejects.toThrow('Busy');
    expect(tokens).toHaveBeenCalledTimes(1);
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('supports gateway error codes and refuses invalid-request SSE errors', async () => {
    global.fetch.mockResolvedValueOnce(sse(['data: {"error":{"code":"upstream_timeout","message":"Timed out"}}\n']))
      .mockResolvedValue(success());
    const outcome = observe(sendMessagesStream(CONFIG, MESSAGES));
    await jest.runAllTimersAsync();
    expect((await outcome).value.content).toBe('Polished text.');
    global.fetch.mockReset().mockResolvedValue(sse(['data: {"error":{"type":"invalid_request_error"}}\n']));
    await expect(sendMessagesStream(CONFIG, MESSAGES)).rejects.toThrow('LLM stream error: unknown');
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test.each([TypeError, SyntaxError])('propagates callback %s errors without retry or swallowing them', async (ErrorType) => {
    global.fetch.mockResolvedValue(sse([contentLine('Text'), 'data: [DONE]\n']));
    await expect(sendMessagesStream(CONFIG, MESSAGES, () => { throw new ErrorType('Callback failed'); }))
      .rejects.toThrow('Callback failed');
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });
});

test.each([false, true])('whole-document dispatch retries only the failed chunk, stream=%s', async (stream) => {
  const counts = new Map();
  global.fetch.mockImplementation(async (_url, { body }) => {
    const key = JSON.parse(body).messages.at(-1).content.includes('Second paragraph') ? 'second' : 'first';
    const count = (counts.get(key) || 0) + 1;
    counts.set(key, count);
    return key === 'second' && count === 1 ? httpError() : success();
  });
  const chunks = ['First paragraph', 'Second paragraph'].map((text, index) => ({ id: `chunk-${index}`, paragraphs: [{ text }] }));
  const log = jest.fn();
  const progress = jest.fn();
  const outcome = observe(processChunksParallel(chunks, {
    config: CONFIG, log, onProgress: progress, onChunkToken: stream ? jest.fn() : undefined,
    promptManager: { getActiveMode: () => 'amendment', getActivePrompt: () => ({ template: '{selection}' }) },
  }));
  await jest.runAllTimersAsync();
  const { value } = await outcome;
  expect(value.map(result => result.status)).toEqual(['fulfilled', 'fulfilled']);
  expect(counts).toEqual(new Map([['first', 1], ['second', 2]]));
  expect(log).toHaveBeenCalledWith(expect.stringContaining('Chunk chunk-1: LLM request failed (HTTP 504); retrying (1/2)'), 'warning');
  expect(progress).toHaveBeenLastCalledWith(expect.objectContaining({ completed: 2, failed: 0, percentComplete: 100 }));
});
