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
    expect(body.temperature).toBe(1);
    expect(body).not.toHaveProperty('reasoning_effort');
  });

  test('sends configured temperature and reasoning effort', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([sseLine('ok'), 'data: [DONE]\n'])
    );

    await sendPromptStream({ ...CONFIG, thinkingLevel: 'high', temperature: 0.6 }, 'prompt', () => {});

    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.temperature).toBe(0.6);
    expect(body.reasoning_effort).toBe('high');
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

describe('sendPromptStream reasoning demux', () => {
  function sseLineFull(delta) {
    return `data: ${JSON.stringify({ choices: [{ delta }] })}\n`;
  }

  test('routes delta.reasoning_content to onReasoning, not the answer stream', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([
        sseLineFull({ reasoning_content: '让我想想' }),
        sseLineFull({ reasoning_content: '……' }),
        sseLineFull({ content: '答案是' }),
        sseLineFull({ content: '42' }),
        'data: [DONE]\n',
      ])
    );

    const content = [];
    const reasoning = [];
    const result = await sendPromptStream(CONFIG, 'prompt', {
      onContent: (t) => content.push(t),
      onReasoning: (t) => reasoning.push(t),
    });

    expect(content).toEqual(['答案是', '42']);
    expect(reasoning).toEqual(['让我想想', '……']);
    expect(result).toBe('答案是42');
  });

  test('demuxes inline <think> blocks across token boundaries', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([
        sseLineFull({ content: '好的<th' }),
        sseLineFull({ content: 'ink>思考一下' }),
        sseLineFull({ content: '这个问题</think>最终回答' }),
        'data: [DONE]\n',
      ])
    );

    const content = [];
    const reasoning = [];
    const result = await sendPromptStream(CONFIG, 'prompt', {
      onContent: (t) => content.push(t),
      onReasoning: (t) => reasoning.push(t),
    });

    expect(content.join('')).toBe('好的最终回答');
    expect(reasoning.join('')).toBe('思考一下这个问题');
    expect(result).toBe('好的最终回答');
  });

  test('non-SSE fallback surfaces message.reasoning_content', async () => {
    global.fetch = jest.fn(async () => ({
      ok: true,
      headers: { get: () => 'application/json' },
      json: async () => ({
        choices: [{ message: { content: '<think>想</think>正文', reasoning_content: '想' } }],
      }),
    }));

    const content = [];
    const reasoning = [];
    const result = await sendPromptStream(CONFIG, 'prompt', {
      onContent: (t) => content.push(t),
      onReasoning: (t) => reasoning.push(t),
    });

    expect(result).toBe('正文');
    expect(reasoning.join('')).toContain('想');
    expect(content.join('')).toBe('正文');
  });

  test('vLLM-style delta.reasoning streams to onReasoning', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([
        sseLineFull({ reasoning: 'thinking ' }),
        sseLineFull({ reasoning: 'hard' }),
        sseLineFull({ content: 'done' }),
        'data: [DONE]\n',
      ])
    );

    const content = [];
    const reasoning = [];
    const result = await sendPromptStream(CONFIG, 'prompt', {
      onContent: (t) => content.push(t),
      onReasoning: (t) => reasoning.push(t),
    });

    expect(content.join('')).toBe('done');
    expect(reasoning.join('')).toBe('thinking hard');
    expect(result).toBe('done');
  });

  test('gateway-style delta.reasoning_details entries stream to onReasoning', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([
        sseLineFull({ reasoning_details: [{ type: 'reasoning.text', text: 'step 1' }] }),
        sseLineFull({ reasoning_details: [{ type: 'reasoning.text', text: ', step 2' }] }),
        sseLineFull({ content: 'answer' }),
        'data: [DONE]\n',
      ])
    );

    const reasoning = [];
    const result = await sendPromptStream(CONFIG, 'prompt', {
      onContent: () => {},
      onReasoning: (t) => reasoning.push(t),
    });

    expect(reasoning.join('')).toBe('step 1, step 2');
    expect(result).toBe('answer');
  });

  test('non-SSE fallback surfaces message.reasoning (vLLM) and reasoning_details', async () => {
    global.fetch = jest.fn(async () => ({
      ok: true,
      headers: { get: () => 'application/json' },
      json: async () => ({
        choices: [{ message: { content: '正文', reasoning: '想清楚了' } }],
      }),
    }));

    const reasoning = [];
    const result = await sendPromptStream(CONFIG, 'prompt', {
      onContent: () => {},
      onReasoning: (t) => reasoning.push(t),
    });

    expect(result).toBe('正文');
    expect(reasoning.join('')).toBe('想清楚了');
  });

  test('streams send the model-specific thinking parameters', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([sseLine('ok'), 'data: [DONE]\n'])
    );

    await sendPromptStream({
      ...CONFIG, provider: 'deepseek', model: 'deepseek-v4-flash',
      thinkingLevel: 'high', temperature: 0.5,
    }, 'prompt', () => {});

    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.thinking).toEqual({ type: 'enabled' });
    expect(body.reasoning_effort).toBe('high');
    expect(body).not.toHaveProperty('temperature');
  });
});

describe('Anthropic streaming transport (claude provider)', () => {
  const { sendMessagesStream } = require('../src/lib/llm-client.js');
  const CLAUDE = {
    url: '/claude', apiKey: 'sk-ant-test', provider: 'claude', model: 'claude-sonnet-4-6',
  };

  /** Anthropic SSE event: an `event:` line plus a `data:` line. */
  function anthropicEvent(type, payload) {
    return `event: ${type}\ndata: ${JSON.stringify({ type, ...payload })}\n\n`;
  }

  test('streams text_delta to content and thinking_delta to reasoning', async () => {
    global.fetch = jest.fn(async () => sseResponse([
      anthropicEvent('message_start', { message: { model: 'claude-sonnet-4-6' } }),
      anthropicEvent('content_block_start', { index: 0, content_block: { type: 'thinking' } }),
      anthropicEvent('content_block_delta', { index: 0, delta: { type: 'thinking_delta', thinking: '思考' } }),
      anthropicEvent('content_block_stop', { index: 0 }),
      anthropicEvent('content_block_start', { index: 1, content_block: { type: 'text' } }),
      anthropicEvent('content_block_delta', { index: 1, delta: { type: 'text_delta', text: '你好' } }),
      anthropicEvent('content_block_delta', { index: 1, delta: { type: 'text_delta', text: '世界' } }),
      anthropicEvent('content_block_stop', { index: 1 }),
      anthropicEvent('message_delta', { delta: { stop_reason: 'end_turn' } }),
      anthropicEvent('message_stop', {}),
    ]));

    const content = [];
    const reasoning = [];
    const result = await sendMessagesStream(CLAUDE, [{ role: 'user', content: 'Hi' }], {
      onContent: (t) => content.push(t),
      onReasoning: (t) => reasoning.push(t),
    });

    expect(result).toEqual({ content: '你好世界', reasoning: '思考' });
    expect(content).toEqual(['你好', '世界']);
    expect(reasoning).toEqual(['思考']);

    // The request went to the Messages endpoint with stream: true.
    const [url, init] = global.fetch.mock.calls[0];
    expect(url).toBe('/claude/v1/messages');
    expect(JSON.parse(init.body).stream).toBe(true);
  });

  test('rejects a stream that ends with stop_reason max_tokens', async () => {
    global.fetch = jest.fn(async () => sseResponse([
      anthropicEvent('content_block_delta', { index: 0, delta: { type: 'text_delta', text: 'half' } }),
      anthropicEvent('message_delta', { delta: { stop_reason: 'max_tokens' } }),
      anthropicEvent('message_stop', {}),
    ]));

    await expect(sendMessagesStream(CLAUDE, [{ role: 'user', content: 'Hi' }], {}))
      .rejects.toThrow(/stop_reason=max_tokens|truncated/);
  });

  test('a closed stream without message_stop or stop_reason is rejected', async () => {
    global.fetch = jest.fn(async () => sseResponse([
      anthropicEvent('content_block_delta', { index: 0, delta: { type: 'text_delta', text: 'partial' } }),
    ]));

    await expect(sendMessagesStream(CLAUDE, [{ role: 'user', content: 'Hi' }], {}))
      .rejects.toThrow(/closed before completion|truncat/i);
  });

  test('non-SSE fallback parses the Anthropic message shape', async () => {
    global.fetch = jest.fn(async () => ({
      ok: true,
      headers: { get: () => 'application/json' },
      json: async () => ({
        stop_reason: 'end_turn',
        content: [
          { type: 'thinking', thinking: '想一下' },
          { type: 'text', text: '正文' },
        ],
      }),
    }));

    const reasoning = [];
    const result = await sendMessagesStream(CLAUDE, [{ role: 'user', content: 'Hi' }], {
      onContent: () => {},
      onReasoning: (t) => reasoning.push(t),
    });

    expect(result.content).toBe('正文');
    expect(reasoning.join('')).toBe('想一下');
  });
});

describe('sendMessagesStream', () => {
  const { sendMessagesStream } = require('../src/lib/llm-client.js');

  test('sends the messages array and returns content plus reasoning', async () => {
    global.fetch = jest.fn(async () =>
      sseResponse([
        `data: ${JSON.stringify({ choices: [{ delta: { reasoning_content: 'r' } }] })}\n`,
        sseLine('out'),
        'data: [DONE]\n',
      ])
    );

    const messages = [
      { role: 'system', content: 'sys' },
      { role: 'user', content: 'user' },
    ];
    const result = await sendMessagesStream(CONFIG, messages, {});

    expect(result).toEqual({ content: 'out', reasoning: 'r' });
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.messages).toEqual(messages);
    expect(body.stream).toBe(true);
  });
});

describe('sendPromptStream idle timeout', () => {
  test('active streaming past the idle window does not abort (clock resets per chunk)', async () => {
    jest.useFakeTimers();
    try {
      const encoder = new TextEncoder();
      const chunks = [sseLine('a'), sseLine('b'), sseLine('c'), 'data: [DONE]\n'];
      let i = 0;
      global.fetch = jest.fn(async () => ({
        ok: true,
        headers: { get: () => 'text/event-stream' },
        body: {
          getReader: () => ({
            read: async () => {
              // Each chunk takes 600ms to arrive: under the 1000ms idle
              // window, but 4 x 600ms = 2400ms exceeds it in total. A
              // total-elapsed timeout would kill this; the idle timeout
              // must not.
              jest.advanceTimersByTime(600);
              return i < chunks.length
                ? { done: false, value: encoder.encode(chunks[i++]) }
                : { done: true, value: undefined };
            },
          }),
        },
      }));

      const tokens = [];
      const result = await sendPromptStream(CONFIG, 'prompt', (t) => tokens.push(t), undefined, undefined, 1000);
      expect(result).toBe('abc');
      expect(tokens).toEqual(['a', 'b', 'c']);
    } finally {
      jest.useRealTimers();
    }
  });

  test('stalled stream aborts with TimeoutError after the idle window', async () => {
    jest.useFakeTimers();
    try {
      const encoder = new TextEncoder();
      const chunks = [sseLine('a'), sseLine('b')];
      let i = 0;
      global.fetch = jest.fn(async () => ({
        ok: true,
        headers: { get: () => 'text/event-stream' },
        body: {
          getReader: () => ({
            read: async () => {
              if (i < chunks.length) {
                return { done: false, value: encoder.encode(chunks[i++]) };
              }
              // Stall: no more data. Advance past the idle window so the
              // abort fires, then emulate the browser rejecting the read.
              jest.advanceTimersByTime(5000);
              throw new DOMException('The operation was aborted.', 'AbortError');
            },
          }),
        },
      }));

      let caught;
      try {
        await sendPromptStream(CONFIG, 'prompt', () => {}, undefined, undefined, 1000);
      } catch (err) {
        caught = err;
      }
      expect(caught).toBeDefined();
      expect(caught.name).toBe('TimeoutError');
      expect(caught.message).toBe('LLM request timed out: no output from the model for 1s');
    } finally {
      jest.useRealTimers();
    }
  });
});

describe('stream / response truncation detection', () => {
  afterEach(() => { delete global.fetch; });

  test('SSE stream closed without [DONE] and without finish_reason is rejected as truncated', async () => {
    // Proxy closed early mid-generation: the tokens so far must NOT be
    // returned as a complete answer (a half-applied amendment is worse
    // than an error).
    global.fetch = jest.fn(async () => sseResponse([sseLine('Partial answ')]));
    await expect(sendPromptStream(CONFIG, 'prompt', () => {}))
      .rejects.toThrow(/closed before completion|truncat/i);
  });

  test('SSE stream with finish_reason (no [DONE]) still resolves', async () => {
    global.fetch = jest.fn(async () => sseResponse([
      `data: ${JSON.stringify({ choices: [{ delta: { content: 'Done' }, finish_reason: 'stop' }] })}\n`,
    ]));
    const result = await sendPromptStream(CONFIG, 'prompt', () => {});
    expect(result).toBe('Done');
  });

  test('non-SSE JSON with finish_reason=length is rejected as truncated', async () => {
    global.fetch = jest.fn(async () => ({
      ok: true,
      headers: { get: () => 'application/json' },
      json: async () => ({ choices: [{ message: { content: 'Half an amend' }, finish_reason: 'length' }] }),
    }));
    await expect(sendPromptStream(CONFIG, 'prompt', () => {}))
      .rejects.toThrow(/truncated|finish_reason/i);
  });

  test('non-streaming sendPrompt rejects finish_reason=length responses', async () => {
    const { sendPrompt } = require('../src/lib/llm-client.js');
    global.fetch = jest.fn(async () => ({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'Truncated ans' }, finish_reason: 'length' }] }),
    }));
    await expect(sendPrompt(CONFIG, 'prompt')).rejects.toThrow(/truncated/i);
  });

  test('non-streaming sendPrompt accepts finish_reason=stop', async () => {
    const { sendPrompt } = require('../src/lib/llm-client.js');
    global.fetch = jest.fn(async () => ({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'Full answer' }, finish_reason: 'stop' }] }),
    }));
    await expect(sendPrompt(CONFIG, 'prompt')).resolves.toBe('Full answer');
  });

  test('SSE stream with finish_reason=length is rejected as truncated', async () => {
    // The full-document pipelines always stream; a length-terminated SSE
    // stream must be refused exactly like the non-streaming paths instead
    // of returning half an amendment as if complete.
    const lengthLine = `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'length' }] })}\n`;
    global.fetch = jest.fn(async () =>
      sseResponse([sseLine('Half an amendment'), lengthLine, 'data: [DONE]\n'])
    );

    await expect(sendPromptStream(CONFIG, 'prompt', () => {}))
      .rejects.toThrow(/finish_reason=length/);
  });

  test('SSE stream with finish_reason=stop completes normally', async () => {
    const stopLine = `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n`;
    global.fetch = jest.fn(async () =>
      sseResponse([sseLine('Complete answer'), stopLine, 'data: [DONE]\n'])
    );

    const tokens = [];
    await expect(sendPromptStream(CONFIG, 'prompt', (t) => tokens.push(t))).resolves.toBe('Complete answer');
    expect(tokens).toEqual(['Complete answer']);
  });
});
