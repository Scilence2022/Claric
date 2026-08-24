/**
 * Tests for sendPromptStream (llm-client.js SSE streaming).
 *
 * Covers:
 *   - SSE parsing: data: lines, chunked delivery, [DONE] terminator
 *   - non-SSE fallback: plain JSON response delivered as a single token
 *   - HTTP error propagation
 *   - abort signal handling
 */

const { sendPromptStream } = require('../src/lib/llm-client.js');

const CONFIG = { url: 'http://llm.local', apiKey: '', model: 'm', apiPath: '/v1' };

/** Builds a mock fetch Response carrying an SSE body delivered in chunks. */
function sseResponse(chunks) {
  const encoder = new TextEncoder();
  let i = 0;
  return {
    ok: true,
    headers: { get: (name) => (name.toLowerCase() === 'content-type' ? 'text/event-stream' : null) },
    body: {
      getReader: () => ({
        read: async () =>
          i < chunks.length
            ? { done: false, value: encoder.encode(chunks[i++]) }
            : { done: true, value: undefined },
      }),
    },
  };
}

/** Builds a mock fetch Response carrying a plain (non-streaming) JSON body. */
function jsonResponse(text) {
  return {
    ok: true,
    headers: { get: () => 'application/json' },
    json: async () => ({ choices: [{ message: { content: text } }] }),
  };
}

function sseLine(delta) {
  return `data: ${JSON.stringify({ choices: [{ delta: { content: delta } }] })}\n`;
}

afterEach(() => {
  jest.restoreAllMocks();
});

describe('sendPromptStream', () => {
  test('streams tokens from SSE data lines until [DONE]', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([sseLine('Hello'), sseLine(' world'), 'data: [DONE]\n', sseLine('ignored')])
    );

    const tokens = [];
    const result = await sendPromptStream(CONFIG, 'prompt', (t) => tokens.push(t));

    expect(tokens).toEqual(['Hello', ' world']);
    expect(result).toBe('Hello world');
    // stream:true was requested
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.stream).toBe(true);
  });

  test('handles SSE lines split across chunks', async () => {
    const payload = sseLine('abc');
    const split = Math.floor(payload.length / 2);
    global.fetch = jest.fn(async () =>
      sseResponse([payload.slice(0, split), payload.slice(split) + 'data: [DONE]\n'])
    );

    const tokens = [];
    const result = await sendPromptStream(CONFIG, 'prompt', (t) => tokens.push(t));

    expect(tokens).toEqual(['abc']);
    expect(result).toBe('abc');
  });

  test('ignores non-data lines and malformed JSON payloads', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([': comment\n\n', 'event: message\n', sseLine('ok'), 'data: {not json}\n', 'data: [DONE]\n'])
    );

    const tokens = [];
    const result = await sendPromptStream(CONFIG, 'prompt', (t) => tokens.push(t));
    expect(tokens).toEqual(['ok']);
    expect(result).toBe('ok');
  });

  test('strips think tags from the assembled response', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([sseLine('<think>reasoning</think>answer'), 'data: [DONE]\n'])
    );

    const result = await sendPromptStream(CONFIG, 'prompt', () => {});
    expect(result).toBe('answer');
  });

  test('falls back to non-streaming when the response is not SSE', async () => {
    global.fetch = jest.fn(async () => jsonResponse('full text'));

    const tokens = [];
    const result = await sendPromptStream(CONFIG, 'prompt', (t) => tokens.push(t));

    expect(result).toBe('full text');
    expect(tokens).toEqual(['full text']);
  });

  test('throws on non-ok HTTP responses', async () => {
    global.fetch = jest.fn(async () => ({ ok: false, status: 500, statusText: 'Server Error' }));
    await expect(sendPromptStream(CONFIG, 'prompt', () => {})).rejects.toThrow('HTTP 500');
  });

  test('rejects immediately with AbortError when the signal is already aborted', async () => {
    global.fetch = jest.fn();
    const controller = new AbortController();
    controller.abort();
    await expect(
      sendPromptStream(CONFIG, 'prompt', () => {}, undefined, controller.signal)
    ).rejects.toMatchObject({ name: 'AbortError' });
    expect(global.fetch).not.toHaveBeenCalled();
  });

  test('external abort during streaming aborts the fetch', async () => {
    const controller = new AbortController();
    global.fetch = jest.fn(async (_url, opts) => {
      // Simulate the fetch rejecting when aborted
      return new Promise((_resolve, reject) => {
        opts.signal.addEventListener('abort', () => {
          reject(new DOMException('The operation was aborted.', 'AbortError'));
        });
      });
    });

    const promise = sendPromptStream(CONFIG, 'prompt', () => {}, undefined, controller.signal);
    controller.abort();
    await expect(promise).rejects.toMatchObject({ name: 'AbortError' });
  });
});
