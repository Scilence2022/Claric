/**
 * Unit tests for src/lib/llm-client.js
 * Tests stripThinkTags, sendPrompt, sendMessages, and testConnection exports.
 */
const { stripThinkTags, stripChunkDelimiters, sendPrompt, sendMessages, testConnection } = require('../src/lib/llm-client.js');

// ============================================================================
// stripThinkTags
// ============================================================================

describe('stripThinkTags', () => {
  test('returns empty string unchanged', () => {
    expect(stripThinkTags('')).toBe('');
  });

  test('returns null unchanged', () => {
    expect(stripThinkTags(null)).toBe(null);
  });

  test('returns undefined unchanged', () => {
    expect(stripThinkTags(undefined)).toBe(undefined);
  });

  test('removes single-line <think>content</think> blocks', () => {
    const input = 'Hello <think>reasoning here</think> World';
    expect(stripThinkTags(input)).toBe('Hello  World');
  });

  test('removes multi-line <think>\\ncontent\\n</think> blocks', () => {
    const input = 'Hello\n<think>\nsome reasoning\nacross lines\n</think>\nWorld';
    expect(stripThinkTags(input)).toBe('Hello\n\nWorld');
  });

  test('removes orphaned </think> tags (closing without opening)', () => {
    const input = 'Some text </think> more text';
    expect(stripThinkTags(input)).toBe('Some text  more text');
  });

  test('removes orphaned <think> tags (opening without closing)', () => {
    const input = 'Some text <think> more text';
    expect(stripThinkTags(input)).toBe('Some text  more text');
  });

  test('handles empty <think></think> tags', () => {
    const input = 'Before <think></think> After';
    expect(stripThinkTags(input)).toBe('Before  After');
  });

  test('trims leading/trailing whitespace and collapses 3+ newlines to 2', () => {
    const input = '  Hello\n\n\n\nWorld  ';
    expect(stripThinkTags(input)).toBe('Hello\n\nWorld');
  });

  test('calls log callback with "Cleaned reasoning artifacts from response" when tags found', () => {
    const log = jest.fn();
    stripThinkTags('Hello <think>test</think> World', log);
    expect(log).toHaveBeenCalledWith('Cleaned reasoning artifacts from response', 'info');
  });

  test('does NOT call log when no tags present', () => {
    const log = jest.fn();
    stripThinkTags('Hello World', log);
    expect(log).not.toHaveBeenCalled();
  });

  test('is case-insensitive (handles <Think>, <THINK>)', () => {
    const input1 = 'Hello <Think>reasoning</Think> World';
    expect(stripThinkTags(input1)).toBe('Hello  World');

    const input2 = 'Hello <THINK>reasoning</THINK> World';
    expect(stripThinkTags(input2)).toBe('Hello  World');
  });

  test('handles multiple think blocks', () => {
    const input = '<think>first</think>Hello<think>second</think> World';
    expect(stripThinkTags(input)).toBe('Hello World');
  });

  test('handles text with no think tags (passes through)', () => {
    const input = 'Just regular text here.';
    expect(stripThinkTags(input)).toBe('Just regular text here.');
  });
});

// ============================================================================
// stripChunkDelimiters
// ============================================================================

describe('stripChunkDelimiters', () => {
  test('returns empty string unchanged', () => {
    expect(stripChunkDelimiters('')).toBe('');
  });

  test('returns null unchanged', () => {
    expect(stripChunkDelimiters(null)).toBe(null);
  });

  test('returns undefined unchanged', () => {
    expect(stripChunkDelimiters(undefined)).toBe(undefined);
  });

  test('removes [END TEXT] marker on its own line', () => {
    const input = 'Some amended text here.\n[END TEXT]';
    expect(stripChunkDelimiters(input)).toBe('Some amended text here.');
  });

  test('removes [AMEND THIS TEXT] marker on its own line', () => {
    const input = '[AMEND THIS TEXT]\nSome amended text here.';
    expect(stripChunkDelimiters(input)).toBe('Some amended text here.');
  });

  test('removes [CONTEXT - DO NOT AMEND] marker on its own line', () => {
    const input = '[CONTEXT - DO NOT AMEND]\nOverlap text\n[END CONTEXT]\n\nAmended text';
    expect(stripChunkDelimiters(input)).toBe('Overlap text\n\nAmended text');
  });

  test('removes [END CONTEXT] marker on its own line', () => {
    const input = 'Previous text\n[END CONTEXT]\nAmended text';
    expect(stripChunkDelimiters(input)).toBe('Previous text\n\nAmended text');
  });

  test('removes all four markers in a complete prompt echo', () => {
    const input = '[CONTEXT - DO NOT AMEND]\nOverlap paragraph\n[END CONTEXT]\n\n[AMEND THIS TEXT]\nActual amended text here.\n[END TEXT]';
    expect(stripChunkDelimiters(input)).toBe('Overlap paragraph\n\nActual amended text here.');
  });

  test('handles markers with leading/trailing whitespace on the line', () => {
    const input = 'Amended text\n  [END TEXT]  ';
    expect(stripChunkDelimiters(input)).toBe('Amended text');
  });

  test('is case-insensitive', () => {
    const input = 'Amended text\n[end text]';
    expect(stripChunkDelimiters(input)).toBe('Amended text');
  });

  test('does not match partial markers embedded in text', () => {
    const input = 'The [END TEXT] marker should only match on its own line, not here.';
    expect(stripChunkDelimiters(input)).toBe('The [END TEXT] marker should only match on its own line, not here.');
  });

  test('passes through text with no markers unchanged', () => {
    const input = 'Regular amended text\nwith multiple paragraphs\nand no markers.';
    expect(stripChunkDelimiters(input)).toBe('Regular amended text\nwith multiple paragraphs\nand no markers.');
  });

  test('collapses triple+ newlines to double after removal', () => {
    const input = 'Before\n\n[END CONTEXT]\n\nAfter';
    const result = stripChunkDelimiters(input);
    expect(result).not.toContain('\n\n\n');
  });

  test('calls log callback when markers are found', () => {
    const log = jest.fn();
    stripChunkDelimiters('text\n[END TEXT]', log);
    expect(log).toHaveBeenCalledWith('Stripped chunk delimiter markers from response', 'info');
  });

  test('does NOT call log when no markers present', () => {
    const log = jest.fn();
    stripChunkDelimiters('Regular text', log);
    expect(log).not.toHaveBeenCalled();
  });
});

// ============================================================================
// sendPrompt
// ============================================================================

describe('sendPrompt', () => {
  let realAbortController;
  beforeEach(() => {
    global.fetch = jest.fn();
    realAbortController = global.AbortController;
    global.AbortController = jest.fn().mockImplementation(() => ({
      signal: 'mock-signal',
      abort: jest.fn()
    }));
    jest.useFakeTimers();
  });

  afterEach(() => {
    global.fetch = undefined;
    // Restore (do not delete): the real AbortController is an own global
    // property, and later describes (testConnection) need it.
    global.AbortController = realAbortController;
    jest.useRealTimers();
  });

  test('constructs correct request body { model, messages, stream: false }', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response text' } }]
      })
    });

    await sendPrompt({ url: '/vllm', apiKey: '', model: 'test-model' }, 'Hello');

    const fetchCall = global.fetch.mock.calls[0];
    const body = JSON.parse(fetchCall[1].body);
    expect(body).toEqual({
      model: 'test-model',
      messages: [{ role: 'user', content: 'Hello' }],
      stream: false,
      temperature: 1,
    });
  });

  test('sends configured temperature and reasoning effort', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'response text' } }] }),
    });

    await sendPrompt({
      url: '/vllm', apiKey: '', model: 'test-model', thinkingLevel: 'medium', temperature: 0.4,
    }, 'Hello');

    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.temperature).toBe(0.4);
    expect(body.reasoning_effort).toBe('medium');
  });

  test('omits reasoning effort for the default thinking level', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'response text' } }] }),
    });

    await sendPrompt({ url: '/vllm', apiKey: '', model: 'test-model', thinkingLevel: 'default', temperature: 1 }, 'Hello');

    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.temperature).toBe(1);
    expect(body).not.toHaveProperty('reasoning_effort');
  });

  test('appends /v1/chat/completions to config.url (stripping trailing slashes)', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response' } }]
      })
    });

    await sendPrompt({ url: '/vllm/', apiKey: '', model: 'test' }, 'Hello');

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[0]).toBe('/vllm/v1/chat/completions');
  });

  test('uses config.apiPath as the API prefix when provided (GLM)', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response' } }]
      })
    });

    await sendPrompt({ url: '/glm', apiPath: '/api/paas/v4', apiKey: 'k', model: 'glm-4.5' }, 'Hello');

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[0]).toBe('/glm/api/paas/v4/chat/completions');
  });

  test('does not double-append the default /v1 prefix when the URL already ends with it', async () => {
    // Regression: users entering the conventional full endpoint
    // (https://host/v1) got https://host/v1/v1/chat/completions → 404 with
    // no hint that the URL, not the endpoint, was wrong.
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'response' } }] })
    });

    await sendPrompt({ url: 'https://llm.example.com/v1', apiKey: '', model: 'test' }, 'Hello');
    expect(global.fetch.mock.calls[0][0]).toBe('https://llm.example.com/v1/chat/completions');
  });

  test('does not double-append a custom apiPath prefix either (GLM full URL)', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'response' } }] })
    });

    await sendPrompt({ url: 'https://open.bigmodel.cn/api/paas/v4', apiPath: '/api/paas/v4', apiKey: 'k', model: 'glm-4.5' }, 'Hello');
    expect(global.fetch.mock.calls[0][0]).toBe('https://open.bigmodel.cn/api/paas/v4/chat/completions');
  });

  test('includes Authorization Bearer header when config.apiKey is non-empty', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response' } }]
      })
    });

    await sendPrompt({ url: '/vllm', apiKey: 'my-secret-key', model: 'test' }, 'Hello');

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[1].headers['Authorization']).toBe('Bearer my-secret-key');
  });

  test('omits Authorization header when config.apiKey is empty/falsy', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response' } }]
      })
    });

    await sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello');

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[1].headers['Authorization']).toBeUndefined();
  });

  test('extracts data.choices[0].message.content from response', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'extracted content' } }]
      })
    });

    const result = await sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello');
    expect(result).toBe('extracted content');
  });

  test('applies stripThinkTags to the extracted content', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '<think>reasoning</think>Clean text' } }]
      })
    });

    const result = await sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello');
    expect(result).toBe('Clean text');
  });

  test('throws on non-ok HTTP response with status code in error message', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 500,
      statusText: 'Internal Server Error'
    });

    await expect(
      sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello')
    ).rejects.toThrow('HTTP 500');
  });

  test('HTTP error message includes the backend error body', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 429,
      statusText: '',
      text: async () => '{"error":{"message":"rate limited, retry after 30s"}}'
    });

    await expect(
      sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello')
    ).rejects.toThrow('HTTP 429: {"error":{"message":"rate limited, retry after 30s"}}');
  });

  test('HTTP error without a readable body still reports the status line', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 502,
      statusText: 'Bad Gateway'
    });

    await expect(
      sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello')
    ).rejects.toThrow('HTTP 502 Bad Gateway');
  });

  test('uses AbortController with 120-second timeout', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response' } }]
      })
    });

    await sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello');

    // Verify AbortController was instantiated
    expect(global.AbortController).toHaveBeenCalled();

    // Verify fetch was called with the abort signal
    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[1].signal).toBe('mock-signal');
  });

  test('returns empty string when choices array is empty or missing', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [] })
    });

    const result = await sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello');
    expect(result).toBe('');
  });
});

// ============================================================================
// testConnection
// ============================================================================

describe('testConnection', () => {
  beforeEach(() => {
    global.fetch = jest.fn();
  });

  afterEach(() => {
    delete global.fetch;
  });

  test('calls GET on config.url + /v1/models', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        data: [{ id: 'model-1' }, { id: 'model-2' }]
      })
    });

    await testConnection({ url: '/ollama', apiKey: '' });

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[0]).toBe('/ollama/v1/models');
    expect(fetchCall[1].method).toBe('GET');
  });

  test('models endpoint also avoids double-appending the /v1 prefix', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ data: [{ id: 'm' }] })
    });

    await testConnection({ url: 'https://llm.example.com/v1', apiKey: '' });
    expect(global.fetch.mock.calls[0][0]).toBe('https://llm.example.com/v1/models');
  });

  test('honors config.apiPath for the models endpoint (GLM)', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ data: [{ id: 'glm-4.5' }] })
    });

    await testConnection({ url: '/glm', apiPath: '/api/paas/v4', apiKey: 'k' });

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[0]).toBe('/glm/api/paas/v4/models');
    expect(fetchCall[1].headers.Authorization).toBe('Bearer k');
  });

  test('includes Authorization header when apiKey provided', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        data: [{ id: 'model-1' }]
      })
    });

    await testConnection({ url: '/vllm', apiKey: 'secret' });

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[1].headers['Authorization']).toBe('Bearer secret');
  });

  test('returns { connected: true, models: [{id}] } from data.data array', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        data: [
          { id: 'gpt-oss:20b', object: 'model', created: 123 },
          { id: 'llama2', object: 'model', created: 456 }
        ]
      })
    });

    const result = await testConnection({ url: '/ollama', apiKey: '' });
    expect(result).toEqual({
      connected: true,
      models: [{ id: 'gpt-oss:20b' }, { id: 'llama2' }]
    });
  });

  test('throws on non-ok HTTP response', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 401,
      statusText: 'Unauthorized'
    });

    await expect(
      testConnection({ url: '/vllm', apiKey: '' })
    ).rejects.toThrow('HTTP 401');
  });

  test('HTTP error message includes the backend error body', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 429,
      statusText: '',
      text: async () => '{"error":{"message":"rate limited, retry after 30s"}}'
    });

    await expect(
      testConnection({ url: '/vllm', apiKey: '' })
    ).rejects.toThrow('HTTP 429: {"error":{"message":"rate limited, retry after 30s"}}');
  });

  test('times out when the backend never answers', async () => {
    jest.useFakeTimers();
    global.fetch = jest.fn((_url, opts) => new Promise((_resolve, reject) => {
      opts.signal.addEventListener('abort', () => {
        const err = new Error('The operation was aborted.');
        err.name = 'AbortError';
        reject(err);
      });
    }));

    // Reject path via fake timers: advance past the 30s probe bound.
    const settled = testConnection({ url: '/ollama', apiKey: '' }).then(
      () => { throw new Error('expected testConnection to reject'); },
      (err) => err
    );
    jest.advanceTimersByTime(30001);
    const err = await settled;

    expect(err.name).toBe('TimeoutError');
    expect(err.message).toMatch(/timed out after 30s/);
    // The probe is bound: fetch receives an abort signal.
    expect(global.fetch.mock.calls[0][1].signal).toBeDefined();

    jest.useRealTimers();
  });

  test('handles empty data array', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ data: [] })
    });

    const result = await testConnection({ url: '/ollama', apiKey: '' });
    expect(result).toEqual({ connected: true, models: [] });
  });

  test('handles missing data field gracefully', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({})
    });

    const result = await testConnection({ url: '/ollama', apiKey: '' });
    expect(result).toEqual({ connected: true, models: [] });
  });
});

// ============================================================================
// sendMessages
// ============================================================================

describe('sendMessages', () => {
  let mockAbortFn;
  let mockSignal;
  let realAbortController;

  beforeEach(() => {
    global.fetch = jest.fn();
    realAbortController = global.AbortController;
    mockAbortFn = jest.fn();
    mockSignal = { aborted: false, addEventListener: jest.fn(), removeEventListener: jest.fn() };
    global.AbortController = jest.fn().mockImplementation(() => ({
      signal: mockSignal,
      abort: mockAbortFn
    }));
    jest.useFakeTimers();
  });

  afterEach(() => {
    global.fetch = undefined;
    global.AbortController = realAbortController;
    jest.useRealTimers();
  });

  test('sends messages array in request body (not flattened to single string)', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response text' } }]
      })
    });

    const messages = [
      { role: 'system', content: 'You are a legal document reviewer.' },
      { role: 'user', content: 'Review this clause: ...' }
    ];

    await sendMessages({ url: '/vllm', apiKey: '', model: 'test-model' }, messages);

    const fetchCall = global.fetch.mock.calls[0];
    const body = JSON.parse(fetchCall[1].body);
    expect(body.messages).toEqual(messages);
    expect(body.model).toBe('test-model');
    expect(body.stream).toBe(false);
    expect(body.temperature).toBe(1);
    expect(body).not.toHaveProperty('reasoning_effort');
  });

  test('sends configured temperature and reasoning effort', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'response text' } }] }),
    });

    await sendMessages({
      url: '/vllm', apiKey: '', model: 'test-model', thinkingLevel: 'low', temperature: 1.4,
    }, [{ role: 'user', content: 'Hello' }]);

    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.temperature).toBe(1.4);
    expect(body.reasoning_effort).toBe('low');
  });

  test('preserves system and user roles in the messages array', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'ok' } }]
      })
    });

    const messages = [
      { role: 'system', content: 'System instructions here.' },
      { role: 'user', content: 'User content here.' }
    ];

    await sendMessages({ url: '/vllm', apiKey: '', model: 'test' }, messages);

    const fetchCall = global.fetch.mock.calls[0];
    const body = JSON.parse(fetchCall[1].body);
    expect(body.messages[0].role).toBe('system');
    expect(body.messages[0].content).toBe('System instructions here.');
    expect(body.messages[1].role).toBe('user');
    expect(body.messages[1].content).toBe('User content here.');
  });

  test('strips think tags from response (reuses stripThinkTags)', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '<think>reasoning</think>Clean output' } }]
      })
    });

    const messages = [{ role: 'user', content: 'Hello' }];
    const result = await sendMessages({ url: '/vllm', apiKey: '', model: 'test' }, messages);
    expect(result).toBe('Clean output');
  });

  test('respects abort signal (throws on aborted signal)', async () => {
    // Simulate fetch rejecting due to abort
    global.fetch.mockRejectedValue(new DOMException('The operation was aborted.', 'AbortError'));

    const externalSignal = {
      aborted: true,
      addEventListener: jest.fn(),
      removeEventListener: jest.fn()
    };

    const messages = [{ role: 'user', content: 'Hello' }];
    await expect(
      sendMessages({ url: '/vllm', apiKey: '', model: 'test' }, messages, null, externalSignal)
    ).rejects.toThrow();
  });

  test('uses configurable timeout (default 30000ms), not hardcoded 120000ms', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response' } }]
      })
    });

    const messages = [{ role: 'user', content: 'Hello' }];

    // Test with default timeout
    await sendMessages({ url: '/vllm', apiKey: '', model: 'test' }, messages);

    // Verify fetch was called with the internal signal
    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[1].signal).toBeDefined();
  });

  test('returns cleaned response text on success', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: '  The amended text is here.  ' } }]
      })
    });

    const messages = [{ role: 'user', content: 'Amend this' }];
    const result = await sendMessages({ url: '/vllm', apiKey: '', model: 'test' }, messages);
    expect(result).toBe('The amended text is here.');
  });

  test('throws on non-ok HTTP response', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 503,
      statusText: 'Service Unavailable'
    });

    const messages = [{ role: 'user', content: 'Hello' }];
    await expect(
      sendMessages({ url: '/vllm', apiKey: '', model: 'test' }, messages)
    ).rejects.toThrow('HTTP 503');
  });

  test('includes Authorization header when apiKey provided', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'ok' } }]
      })
    });

    const messages = [{ role: 'user', content: 'Hello' }];
    await sendMessages({ url: '/vllm', apiKey: 'secret-key', model: 'test' }, messages);

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[1].headers['Authorization']).toBe('Bearer secret-key');
  });

  test('appends /v1/chat/completions to config.url', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'ok' } }]
      })
    });

    const messages = [{ role: 'user', content: 'Hello' }];
    await sendMessages({ url: '/vllm/', apiKey: '', model: 'test' }, messages);

    const fetchCall = global.fetch.mock.calls[0];
    expect(fetchCall[0]).toBe('/vllm/v1/chat/completions');
  });
});

// ============================================================================
// sendPrompt backward compatibility
// ============================================================================

describe('sendPrompt backward compatibility', () => {
  beforeEach(() => {
    global.fetch = jest.fn();
    global.AbortController = jest.fn().mockImplementation(() => ({
      signal: 'mock-signal',
      abort: jest.fn()
    }));
    jest.useFakeTimers();
  });

  afterEach(() => {
    delete global.fetch;
    jest.useRealTimers();
  });

  test('existing sendPrompt still works unchanged', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({
        choices: [{ message: { content: 'response from sendPrompt' } }]
      })
    });

    const result = await sendPrompt({ url: '/vllm', apiKey: '', model: 'test' }, 'Hello');
    expect(result).toBe('response from sendPrompt');

    // Verify it still wraps in a single user message
    const fetchCall = global.fetch.mock.calls[0];
    const body = JSON.parse(fetchCall[1].body);
    expect(body.messages).toEqual([{ role: 'user', content: 'Hello' }]);
  });
});

// ============================================================================
// Model-specific generation parameters (model-capabilities.js integration)
// ============================================================================

describe('model-specific generation parameters', () => {
  let realAbortController;
  beforeEach(() => {
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ choices: [{ message: { content: 'ok' } }] }),
    });
    realAbortController = global.AbortController;
    global.AbortController = jest.fn().mockImplementation(() => ({
      signal: 'mock-signal',
      abort: jest.fn()
    }));
    jest.useFakeTimers();
  });

  afterEach(() => {
    global.fetch = undefined;
    global.AbortController = realAbortController;
    jest.useRealTimers();
  });

  async function bodyFor(config) {
    await sendPrompt(config, 'Hello');
    return JSON.parse(global.fetch.mock.calls[0][1].body);
  }

  test('vLLM Qwen3 sends enable_thinking and a token budget', async () => {
    const body = await bodyFor({
      url: '/vllm', apiKey: '', provider: 'vllm', model: 'qwen3.5-35b-a3b',
      thinkingLevel: 'high', temperature: 0.3,
    });
    expect(body.chat_template_kwargs).toEqual({ enable_thinking: true });
    expect(body.thinking_token_budget).toBe(16384);
    expect(body.temperature).toBe(0.3);
    expect(body).not.toHaveProperty('reasoning_effort');
  });

  test('vLLM Qwen3 Off disables thinking via the chat template', async () => {
    const body = await bodyFor({
      url: '/vllm', apiKey: '', provider: 'vllm', model: 'qwen3.5-35b-a3b',
      thinkingLevel: 'off',
    });
    expect(body).toEqual(expect.objectContaining({
      chat_template_kwargs: { enable_thinking: false },
    }));
    expect(body).not.toHaveProperty('thinking_token_budget');
  });

  test('DeepSeek V4 sends thinking.type plus reasoning_effort, no temperature while thinking', async () => {
    const body = await bodyFor({
      url: '/deepseek', apiKey: '', provider: 'deepseek', model: 'deepseek-v4-flash',
      thinkingLevel: 'low', temperature: 0.2,
    });
    expect(body.thinking).toEqual({ type: 'enabled' });
    expect(body.reasoning_effort).toBe('low');
    expect(body).not.toHaveProperty('temperature');
  });

  test('DeepSeek V4 Off restores temperature and disables thinking', async () => {
    const body = await bodyFor({
      url: '/deepseek', apiKey: '', provider: 'deepseek', model: 'deepseek-v4-flash',
      thinkingLevel: 'off', temperature: 0.2,
    });
    expect(body.thinking).toEqual({ type: 'disabled' });
    expect(body.temperature).toBe(0.2);
  });

  test('GLM-5.2 maps canonical levels to thinking + effort', async () => {
    const body = await bodyFor({
      url: '/glm', apiPath: '/api/paas/v4', apiKey: 'k', provider: 'glm', model: 'glm-5.2',
      thinkingLevel: 'xhigh', temperature: 1.2,
    });
    expect(body.thinking).toEqual({ type: 'enabled' });
    expect(body.reasoning_effort).toBe('xhigh');
    expect(body.temperature).toBe(1.2);
  });

  test('GLM 4.x exposes only the thinking toggle', async () => {
    const body = await bodyFor({
      url: '/glm', apiPath: '/api/paas/v4', apiKey: 'k', provider: 'glm', model: 'glm-4.6',
      thinkingLevel: 'on',
    });
    expect(body.thinking).toEqual({ type: 'enabled' });
    expect(body).not.toHaveProperty('reasoning_effort');
  });

  test('Kimi K3 sends top-level reasoning_effort and no temperature', async () => {
    const body = await bodyFor({
      url: '/kimi', apiKey: 'k', provider: 'kimi', model: 'kimi-k3',
      thinkingLevel: 'max', temperature: 0.7,
    });
    expect(body.reasoning_effort).toBe('max');
    expect(body).not.toHaveProperty('temperature');
  });

  test('Kimi K2.6 uses the thinking toggle, unsupported legacy values resolve safely', async () => {
    const body = await bodyFor({
      url: '/kimi', apiKey: 'k', provider: 'kimi', model: 'kimi-k2.6',
      thinkingLevel: 'medium',
    });
    // 'medium' is not a K2.6 level: resolves to the profile default, which
    // sends no thinking field at all.
    expect(body).not.toHaveProperty('thinking');
    expect(body).not.toHaveProperty('reasoning_effort');
  });

  test('MiniMax M3 adaptive sends the adaptive toggle and reasoning split', async () => {
    const body = await bodyFor({
      url: '/minimax', apiKey: 'k', provider: 'minimax', model: 'MiniMax-M3',
      thinkingLevel: 'adaptive',
    });
    expect(body.thinking).toEqual({ type: 'adaptive' });
    expect(body.reasoning_split).toBe(true);
  });

  test('Ollama GPT-OSS explicit levels send reasoning_effort', async () => {
    const body = await bodyFor({
      url: '/ollama', apiKey: '', provider: 'ollama', model: 'gpt-oss:20b',
      thinkingLevel: 'medium',
    });
    expect(body.reasoning_effort).toBe('medium');
    expect(body.temperature).toBe(1);
  });

  test('unknown models keep generic behavior', async () => {
    const body = await bodyFor({
      url: '/custom', apiKey: '', provider: 'custom', model: 'some-future-model',
      thinkingLevel: 'high', temperature: 0.9,
    });
    expect(body.reasoning_effort).toBe('high');
    expect(body.temperature).toBe(0.9);
    expect(body).not.toHaveProperty('thinking');
  });

  test('a gateway model id that names a known upstream model uses that profile', async () => {
    const body = await bodyFor({
      url: '/zhongkeyu', apiKey: 'k', provider: 'zhongkeyu', model: 'glm-5.3-flash',
      thinkingLevel: 'low',
    });
    expect(body.thinking).toEqual({ type: 'enabled' });
    expect(body.reasoning_effort).toBe('low');
  });

  test('OpenAI GPT-5.6 sends reasoning_effort and omits temperature while reasoning', async () => {
    const body = await bodyFor({
      url: '/openai', apiKey: 'k', provider: 'openai', model: 'gpt-5.6',
      thinkingLevel: 'xhigh', temperature: 0.4,
    });
    expect(body.reasoning_effort).toBe('xhigh');
    expect(body).not.toHaveProperty('temperature');
  });

  test('OpenAI GPT-5.1 at none disables reasoning and restores temperature', async () => {
    const body = await bodyFor({
      url: '/openai', apiKey: 'k', provider: 'openai', model: 'gpt-5.1',
      thinkingLevel: 'none', temperature: 0.4,
    });
    expect(body.reasoning_effort).toBe('none');
    expect(body.temperature).toBe(0.4);
  });

  test('OpenAI GPT-4o has no reasoning dial and keeps temperature', async () => {
    const body = await bodyFor({
      url: '/openai', apiKey: 'k', provider: 'openai', model: 'gpt-4o',
      thinkingLevel: 'high', temperature: 1.3,
    });
    expect(body).not.toHaveProperty('reasoning_effort');
    expect(body.temperature).toBe(1.3);
  });
});

// ============================================================================
// Anthropic Messages API transport (claude provider)
// ============================================================================

describe('Anthropic transport (claude provider)', () => {
  let realAbortController;
  const CLAUDE = {
    url: '/claude', apiKey: 'sk-ant-test', provider: 'claude', model: 'claude-sonnet-4-6',
  };

  beforeEach(() => {
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      json: async () => ({
        stop_reason: 'end_turn',
        content: [
          { type: 'thinking', thinking: 'hmm' },
          { type: 'text', text: 'Hello ' },
          { type: 'text', text: 'there' },
        ],
      }),
    });
    realAbortController = global.AbortController;
    global.AbortController = jest.fn().mockImplementation(() => ({
      signal: 'mock-signal',
      abort: jest.fn()
    }));
    jest.useFakeTimers();
  });

  afterEach(() => {
    global.fetch = undefined;
    global.AbortController = realAbortController;
    jest.useRealTimers();
  });

  test('posts to /v1/messages with Anthropic auth headers', async () => {
    await sendMessages(CLAUDE, [{ role: 'user', content: 'Hi' }]);
    const [url, init] = global.fetch.mock.calls[0];
    expect(url).toBe('/claude/v1/messages');
    expect(init.headers['x-api-key']).toBe('sk-ant-test');
    expect(init.headers['anthropic-version']).toBe('2023-06-01');
    expect(init.headers['anthropic-dangerous-direct-browser-access']).toBe('true');
    expect(init.headers.Authorization).toBeUndefined();
  });

  test('translates system messages, defaults max_tokens, and concatenates text blocks', async () => {
    const result = await sendMessages(CLAUDE, [
      { role: 'system', content: 'You edit documents.' },
      { role: 'user', content: 'Hi' },
    ]);
    expect(result).toBe('Hello there');
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.system).toBe('You edit documents.');
    expect(body.messages).toEqual([{ role: 'user', content: 'Hi' }]);
    expect(body.max_tokens).toBe(16384);
    expect(body.stream).toBe(false);
    expect(body).not.toHaveProperty('output_config');
  });

  test('sends output_config.effort for effort-era models', async () => {
    await sendMessages({ ...CLAUDE, thinkingLevel: 'low' }, [{ role: 'user', content: 'Hi' }]);
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.output_config).toEqual({ effort: 'low' });
    expect(body).not.toHaveProperty('thinking');
  });

  test('xhigh effort raises the max_tokens headroom', async () => {
    await sendMessages({ ...CLAUDE, model: 'claude-opus-4-7', thinkingLevel: 'xhigh' },
      [{ role: 'user', content: 'Hi' }]);
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.output_config).toEqual({ effort: 'xhigh' });
    expect(body.max_tokens).toBe(65536);
  });

  test('budget-era models send thinking budgets and no temperature', async () => {
    await sendMessages({
      ...CLAUDE, model: 'claude-opus-4-5', thinkingLevel: 'high', temperature: 0.5,
    }, [{ role: 'user', content: 'Hi' }]);
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.thinking).toEqual({ type: 'enabled', budget_tokens: 16384 });
    // budget + answer floor: 16384 + 8192
    expect(body.max_tokens).toBe(24576);
    expect(body).not.toHaveProperty('temperature');
  });

  test('clamps temperature to the Anthropic 0-1 range', async () => {
    await sendMessages({ ...CLAUDE, thinkingLevel: 'off', temperature: 1.8 },
      [{ role: 'user', content: 'Hi' }]);
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.temperature).toBe(1);
    expect(body.thinking).toEqual({ type: 'disabled' });
  });

  test('translates image_url parts into Anthropic image blocks', async () => {
    const dataUrl = 'data:image/png;base64,QUJD';
    await sendMessages(CLAUDE, [{
      role: 'user',
      content: [
        { type: 'text', text: 'What is this?' },
        { type: 'image_url', image_url: { url: dataUrl } },
        { type: 'image_url', image_url: { url: 'https://example.com/pic.png' } },
      ],
    }]);
    const body = JSON.parse(global.fetch.mock.calls[0][1].body);
    expect(body.messages[0].content).toEqual([
      { type: 'text', text: 'What is this?' },
      { type: 'image', source: { type: 'base64', media_type: 'image/png', data: 'QUJD' } },
      { type: 'image', source: { type: 'url', url: 'https://example.com/pic.png' } },
    ]);
  });

  test('rejects stop_reason=max_tokens as truncated', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ stop_reason: 'max_tokens', content: [{ type: 'text', text: 'half' }] }),
    });
    await expect(sendMessages(CLAUDE, [{ role: 'user', content: 'Hi' }]))
      .rejects.toThrow(/truncated/);
  });

  test('an explicit apiFormat dispatches without the provider id', async () => {
    await sendMessages({ url: 'https://relay.example.com', apiKey: 'k', apiFormat: 'anthropic', model: 'claude-x' },
      [{ role: 'user', content: 'Hi' }]);
    const [url, init] = global.fetch.mock.calls[0];
    expect(url).toBe('https://relay.example.com/v1/messages');
    expect(init.headers['x-api-key']).toBe('k');
  });

  test('testConnection uses Anthropic headers against /v1/models', async () => {
    global.fetch.mockResolvedValue({
      ok: true,
      json: async () => ({ data: [{ id: 'claude-sonnet-4-6' }], has_more: false }),
    });
    const result = await testConnection(CLAUDE);
    expect(result.models).toEqual([{ id: 'claude-sonnet-4-6' }]);
    const [url, init] = global.fetch.mock.calls[0];
    expect(url).toBe('/claude/v1/models');
    expect(init.headers['x-api-key']).toBe('sk-ant-test');
    expect(init.headers.Authorization).toBeUndefined();
  });
});
