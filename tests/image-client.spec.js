/**
 * Image-generation client tests.
 *
 * generateImage exists to absorb one specific mess: text-to-image providers
 * answer the same request in three different shapes (inline b64_json, a data
 * URL, or a hosted file URL), while Word's insertInlinePictureFromBase64
 * accepts exactly one (raster base64). Every branch of that normalization is
 * pinned here, because a shape the client mishandles surfaces to the user as a
 * silently blank illustration rather than an error.
 *
 * The rest covers the parts that cost a user a real round trip when wrong:
 * request shaping (response_format only where the provider honors it, size
 * omitted when 'auto', Authorization only when a key exists), idempotent URL
 * joining (a pasted `.../v1` must not become `/v1/v1/...`), actionable errors
 * for unconfigured or failing providers, and abort handling — including
 * testImageConnection's rule that it converts failures into a report but must
 * still let cancellation propagate.
 *
 * Runs under the default node environment: the client needs fetch (mocked) and
 * btoa, which node provides globally. No DOM required.
 */

const {
  generateImage,
  testImageConnection,
  isAllowedImageEndpoint,
  isAllowedImageDownloadUrl,
  sanitizeImageErrorMessage,
} = require('../src/lib/image-client.js');

/** A minimal config that passes the guard clauses. */
const CONFIG = { url: '/openai', apiKey: '', model: 'gpt-image-1', provider: 'openai' };

/** Resolves a mock JSON response like fetch would. */
function jsonResponse(payload) {
  return { ok: true, json: async () => payload };
}

/** Reads the parsed request body of the Nth fetch call. */
function bodyOf(callIndex = 0) {
  return JSON.parse(global.fetch.mock.calls[callIndex][1].body);
}

beforeEach(() => {
  global.fetch = jest.fn();
});

afterEach(() => {
  global.fetch = undefined;
  jest.clearAllMocks();
});

// ============================================================================
// Response shape normalization
// ============================================================================

describe('generateImage response shapes', () => {
  test('returns inline b64_json as-is', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'QUJD' }] }));

    const result = await generateImage(CONFIG, 'a grey circle');

    expect(result.base64).toBe('QUJD');
    expect(result.model).toBe('gpt-image-1');
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('reports the configured size, defaulting when config carries none', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'QUJD' }] }));

    const sized = await generateImage({ ...CONFIG, size: '1536x1024' }, 'x');
    expect(sized.size).toBe('1536x1024');

    const unsized = await generateImage(CONFIG, 'x');
    expect(unsized.size).toBe('1024x1024');
  });

  test('unwraps a data URL without issuing a second request', async () => {
    global.fetch.mockResolvedValue(jsonResponse({
      data: [{ url: 'data:image/png;base64,SGVsbG8=' }],
    }));

    const result = await generateImage(CONFIG, 'a grey circle');

    expect(result.base64).toBe('SGVsbG8=');
    // The whole point of the data-URL branch: no network round trip for bytes.
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('fetches a hosted URL and converts the bytes to base64', async () => {
    const bytes = new Uint8Array([0x48, 0x65, 0x6c, 0x6c, 0x6f]); // "Hello"
    global.fetch
      .mockResolvedValueOnce(jsonResponse({ data: [{ url: 'https://cdn.example.com/img.png' }] }))
      .mockResolvedValueOnce({ ok: true, arrayBuffer: async () => bytes.buffer });

    const result = await generateImage(CONFIG, 'a grey circle');

    expect(result.base64).toBe(btoa('Hello'));
    expect(global.fetch).toHaveBeenCalledTimes(2);
    expect(global.fetch.mock.calls[1][0]).toBe('https://cdn.example.com/img.png');
  });

  test('prefers b64_json over a URL when both are present', async () => {
    global.fetch.mockResolvedValue(jsonResponse({
      data: [{ b64_json: 'QUJD', url: 'https://cdn.example.com/img.png' }],
    }));

    const result = await generateImage(CONFIG, 'a grey circle');

    expect(result.base64).toBe('QUJD');
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('a failed bytes download reports the status and the static-host hint', async () => {
    global.fetch
      .mockResolvedValueOnce(jsonResponse({ data: [{ url: 'https://cdn.example.com/img.png' }] }))
      .mockResolvedValueOnce({ ok: false, status: 403 });

    const err = await generateImage(CONFIG, 'a grey circle').catch((e) => e);
    expect(err.message).toMatch(/HTTP 403/);
    expect(err.message).toMatch(/static host/);
  });
});

// ============================================================================
// Unusable responses
// ============================================================================

describe('generateImage unusable responses', () => {
  test('throws when the item carries neither b64_json nor url', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ revised_prompt: 'a grey circle' }] }));

    await expect(generateImage(CONFIG, 'x'))
      .rejects.toThrow('The image provider returned neither base64 data nor a URL.');
  });

  test('throws when the data array is empty or missing', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [] }));
    await expect(generateImage(CONFIG, 'x'))
      .rejects.toThrow('The image provider returned no image data.');

    global.fetch.mockResolvedValue(jsonResponse({}));
    await expect(generateImage(CONFIG, 'x'))
      .rejects.toThrow('The image provider returned no image data.');
  });

  test('an empty b64_json string is not accepted as an image', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: '' }] }));

    await expect(generateImage(CONFIG, 'x'))
      .rejects.toThrow('The image provider returned neither base64 data nor a URL.');
  });
});

// ============================================================================
// Configuration guards
// ============================================================================

describe('generateImage configuration guards', () => {
  test('a missing endpoint points the user at Settings and skips fetch', async () => {
    await expect(generateImage({ model: 'gpt-image-1' }, 'x'))
      .rejects.toThrow('No image provider endpoint configured. Set one in Settings → Image Generation.');
    await expect(generateImage(null, 'x')).rejects.toThrow(/Settings → Image Generation/);
    expect(global.fetch).not.toHaveBeenCalled();
  });

  test('a missing model points the user at Settings and skips fetch', async () => {
    await expect(generateImage({ url: '/openai' }, 'x'))
      .rejects.toThrow('No image model configured. Set one in Settings → Image Generation.');
    expect(global.fetch).not.toHaveBeenCalled();
  });

  test('an empty or non-string prompt is rejected before any request', async () => {
    await expect(generateImage(CONFIG, '')).rejects.toThrow('An image prompt is required.');
    await expect(generateImage(CONFIG, '   \n ')).rejects.toThrow('An image prompt is required.');
    await expect(generateImage(CONFIG, null)).rejects.toThrow('An image prompt is required.');
    expect(global.fetch).not.toHaveBeenCalled();
  });
});

describe('image URL validation and error redaction', () => {
  test.each([
    ['/openai', true],
    ['./proxy', true],
    ['https://api.example.com/v1', true],
    ['http://localhost:11434', true],
    ['http://127.0.0.1:8026', true],
    ['http://[::1]:8026', true],
    ['http://remote.example/v1', false],
    ['ftp://remote.example/image', false],
    ['javascript:alert(1)', false],
    ['file:///tmp/image.png', false],
    ['data:image/png;base64,QUJD', false],
    ['//remote.example/image.png', false],
    ['/\\remote.example/image.png', false],
    ['https://user:pass@example.com/image.png', false],
  ])('classifies endpoint %p as %p', (url, expected) => {
    expect(isAllowedImageEndpoint(url)).toBe(expected);
  });

  test.each([
    ['https://cdn.example.com/image.png', true],
    ['http://localhost:9000/image.png', true],
    ['http://127.0.0.1/image.png', true],
    ['http://[::1]/image.png', true],
    ['/same-origin/image.png', false],
    ['http://remote.example/image.png', false],
    ['data:image/png;base64,QUJD', false],
    ['javascript:alert(1)', false],
  ])('classifies download URL %p as %p', (url, expected) => {
    expect(isAllowedImageDownloadUrl(url)).toBe(expected);
  });

  test('rejects an unsafe endpoint before fetch', async () => {
    await expect(generateImage({
      url: 'http://remote.example/images?token=img-secret',
      model: 'gpt-image-1',
      apiKey: 'img-secret',
    }, 'x')).rejects.toThrow(/HTTPS.*localhost HTTP.*relative proxy path/);
    expect(global.fetch).not.toHaveBeenCalled();
  });

  test('does not fetch a remote cleartext image URL returned by a provider', async () => {
    global.fetch.mockResolvedValue(jsonResponse({
      data: [{ url: 'http://remote.example/image.png?token=img-secret' }],
    }));

    const err = await generateImage(CONFIG, 'x').catch((e) => e);
    expect(err.message).toBe('The image provider returned an unsupported image URL.');
    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(err.message).not.toContain('remote.example');
    expect(err.message).not.toContain('img-secret');
  });

  test('a hosted-image network error does not reveal its URL or token', async () => {
    const hostedUrl = 'https://cdn.example.com/img.png?token=download-secret';
    global.fetch
      .mockResolvedValueOnce(jsonResponse({ data: [{ url: hostedUrl }] }))
      .mockRejectedValueOnce(new Error(`fetch failed for ${hostedUrl}`));

    const err = await generateImage(CONFIG, 'x').catch((e) => e);
    expect(err.message).toBe('The image provider image download failed.');
    expect(err.message).not.toContain(hostedUrl);
    expect(err.message).not.toContain('download-secret');
  });

  test('allows a loopback HTTP image download', async () => {
    const bytes = new Uint8Array([0x50, 0x4e, 0x47]);
    global.fetch
      .mockResolvedValueOnce(jsonResponse({ data: [{ url: 'http://127.0.0.1:9000/image.png' }] }))
      .mockResolvedValueOnce({ ok: true, arrayBuffer: async () => bytes.buffer });

    await expect(generateImage(CONFIG, 'x')).resolves.toHaveProperty('base64', btoa('PNG'));
    expect(global.fetch.mock.calls[1][0]).toBe('http://127.0.0.1:9000/image.png');
  });

  test('redacts OpenAI error URLs, query keys, and bearer tokens', async () => {
    const token = 'sk-image-secret';
    global.fetch.mockResolvedValue({
      ok: false,
      status: 401,
      statusText: 'Unauthorized',
      text: async () => `Authorization: Bearer ${token}; see https://api.example.com/v1?api_key=${token}`,
    });

    const err = await generateImage({ ...CONFIG, apiKey: token }, 'x').catch((e) => e);
    expect(err.message).toContain('HTTP 401');
    expect(err.message).not.toContain(token);
    expect(err.message).not.toContain('https://api.example.com');
  });

  test('redacts MiniMax provider status details', async () => {
    const token = 'minimax-image-secret';
    global.fetch.mockResolvedValue(jsonResponse({
      base_resp: {
        status_code: 1001,
        status_msg: `x-api-key: ${token}; retry https://api.minimax.io/v1?token=${token}`,
      },
    }));

    const err = await generateImage({
      url: '/minimax',
      model: 'image-01',
      provider: 'minimax',
      apiFormat: 'minimax-images',
      apiKey: token,
    }, 'x').catch((e) => e);
    expect(err.message).toContain('MiniMax image request failed (1001)');
    expect(err.message).not.toContain(token);
    expect(err.message).not.toContain('https://api.minimax.io');
  });

  test('sanitizes standalone sensitive fields and absolute URLs', () => {
    const message = sanitizeImageErrorMessage(
      'x-api-key: secret; {"api_key":"secret"}; https://example.com/path?token=secret',
      ['secret'],
    );
    expect(message).not.toContain('secret');
    expect(message).not.toContain('https://example.com');
    expect(message).toContain('[redacted]');
  });
});

// ============================================================================
// HTTP failures
// ============================================================================

describe('generateImage HTTP failures', () => {
  test('error message carries the status code and the upstream body', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 400,
      statusText: 'Bad Request',
      text: async () => '{"error":{"message":"content policy violation"}}',
    });

    await expect(generateImage(CONFIG, 'x'))
      .rejects.toThrow('Image request failed: HTTP 400 Bad Request — {"error":{"message":"content policy violation"}}');
  });

  test('an unreadable body still yields the status line', async () => {
    global.fetch.mockResolvedValue({
      ok: false,
      status: 502,
      statusText: 'Bad Gateway',
      text: async () => { throw new Error('body already consumed'); },
    });

    await expect(generateImage(CONFIG, 'x'))
      .rejects.toThrow('Image request failed: HTTP 502 Bad Gateway');
  });

  test('a flooding HTML error page is truncated to 300 characters', async () => {
    const huge = 'E'.repeat(5000);
    global.fetch.mockResolvedValue({
      ok: false, status: 500, statusText: '', text: async () => huge,
    });

    const err = await generateImage(CONFIG, 'x').catch((e) => e);
    expect(err.message).toContain('E'.repeat(300));
    expect(err.message).not.toContain('E'.repeat(301));
  });
});

// ============================================================================
// Request shaping
// ============================================================================

describe('generateImage request shaping', () => {
  beforeEach(() => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'QUJD' }] }));
  });

  test('omits response_format for an OpenAI GPT Image model', async () => {
    await generateImage({ url: '/openai', model: 'gpt-image-1', provider: 'openai' }, 'a circle');

    // GPT Image models always return base64 and reject response_format.
    expect(bodyOf()).toEqual({
      model: 'gpt-image-1',
      prompt: 'a circle',
      n: 1,
    });
  });

  test('requests b64_json for an OpenAI DALL-E model', async () => {
    await generateImage({ url: '/openai', model: 'dall-e-3', provider: 'openai' }, 'a circle');

    expect(bodyOf()).toEqual({
      model: 'dall-e-3',
      prompt: 'a circle',
      n: 1,
      response_format: 'b64_json',
    });
  });

  test('omits response_format for a provider that rejects it (glm preset)', async () => {
    await generateImage({ url: '/glm', apiPath: '/api/paas/v4', model: 'cogview-4', provider: 'glm' }, 'a circle');

    expect(bodyOf()).not.toHaveProperty('response_format');
    expect(bodyOf().model).toBe('cogview-4');
  });

  test('uses MiniMax image_generation endpoint and native request fields', async () => {
    global.fetch.mockResolvedValue(jsonResponse({
      data: { image_base64: ['QUJD'] },
      base_resp: { status_code: 0, status_msg: 'success' },
    }));

    await generateImage({
      url: '/minimax', apiPath: '/v1', model: 'image-01', provider: 'minimax',
      apiFormat: 'minimax-images', size: '1536x1024',
    }, 'a circle');

    expect(global.fetch.mock.calls[0][0]).toBe('/minimax/v1/image_generation');
    expect(bodyOf()).toEqual({
      model: 'image-01',
      prompt: 'a circle',
      aspect_ratio: '3:2',
      response_format: 'base64',
      n: 1,
      prompt_optimizer: true,
    });
  });

  test('extracts MiniMax inline base64 and hosted URL response shapes', async () => {
    global.fetch
      .mockResolvedValueOnce(jsonResponse({
        data: { image_base64: ['INLINE_IMAGE'] },
        base_resp: { status_code: 0 },
      }))
      .mockResolvedValueOnce(jsonResponse({
        data: { image_urls: ['https://cdn.example.com/minimax.png'] },
        base_resp: { status_code: 0 },
      }))
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: async () => new Uint8Array([0x50, 0x4e, 0x47]).buffer,
      });

    const config = {
      url: '/minimax', apiPath: '/v1', model: 'image-01', provider: 'minimax',
      apiFormat: 'minimax-images',
    };
    await expect(generateImage(config, 'inline')).resolves.toHaveProperty('base64', 'INLINE_IMAGE');
    await expect(generateImage(config, 'hosted')).resolves.toHaveProperty('base64', btoa('PNG'));
    expect(global.fetch.mock.calls[1][0]).toBe('/minimax/v1/image_generation');
    expect(global.fetch.mock.calls[2][0]).toBe('https://cdn.example.com/minimax.png');
  });

  test('truncates MiniMax prompts to the documented 1500-character limit', async () => {
    await generateImage({
      url: '/minimax', model: 'image-01', provider: 'minimax', apiFormat: 'minimax-images',
    }, 'x'.repeat(2000));

    const sentPrompt = bodyOf().prompt;
    expect(sentPrompt).toHaveLength(1500);
    expect(sentPrompt).toMatch(/\[Document context truncated\]$/);
  });

  test('surfaces MiniMax base_resp errors without accepting image data', async () => {
    global.fetch.mockResolvedValue(jsonResponse({
      base_resp: { status_code: 1001, status_msg: 'invalid prompt' },
      data: { image_base64: ['SHOULD_NOT_BE_USED'] },
    }));

    await expect(generateImage({
      url: '/minimax', model: 'image-01', provider: 'minimax', apiFormat: 'minimax-images',
    }, 'a circle')).rejects.toThrow('MiniMax image request failed (1001): invalid prompt');
  });

  test('an explicit wantsB64 flag overrides an OpenAI-compatible relay policy', async () => {
    await generateImage({
      url: '/relay', model: 'relay-model', provider: 'custom', wantsB64: true,
    }, 'a circle');
    expect(bodyOf().response_format).toBe('b64_json');

    await generateImage({
      url: '/relay', model: 'relay-model', provider: 'custom', wantsB64: false,
    }, 'a circle');
    expect(bodyOf(1)).not.toHaveProperty('response_format');
  });

  test('an unknown provider defaults to sending no response_format', async () => {
    await generateImage({ url: '/relay', model: 'some-model', provider: 'not-a-provider' }, 'a circle');
    expect(bodyOf()).not.toHaveProperty('response_format');
  });

  test("sends a concrete size but omits the field for 'auto'", async () => {
    await generateImage({ ...CONFIG, size: '1024x1536' }, 'a circle');
    expect(bodyOf().size).toBe('1024x1536');

    await generateImage({ ...CONFIG, size: 'auto' }, 'a circle');
    expect(bodyOf(1)).not.toHaveProperty('size');

    await generateImage(CONFIG, 'a circle');
    expect(bodyOf(2)).not.toHaveProperty('size');
  });

  test('trims the prompt before sending it', async () => {
    await generateImage(CONFIG, '  a grey circle  ');
    expect(bodyOf().prompt).toBe('a grey circle');
  });

  test('includes Authorization: Bearer only when an apiKey is set', async () => {
    await generateImage({ ...CONFIG, apiKey: 'sk-test-123' }, 'a circle');
    expect(global.fetch.mock.calls[0][1].headers['Authorization']).toBe('Bearer sk-test-123');

    await generateImage({ ...CONFIG, apiKey: '' }, 'a circle');
    expect(global.fetch.mock.calls[1][1].headers['Authorization']).toBeUndefined();
  });

  test('posts JSON', async () => {
    await generateImage(CONFIG, 'a circle');

    const [, options] = global.fetch.mock.calls[0];
    expect(options.method).toBe('POST');
    expect(options.headers['Content-Type']).toBe('application/json');
  });
});

// ============================================================================
// URL joining
// ============================================================================

describe('generateImage URL joining', () => {
  beforeEach(() => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'QUJD' }] }));
  });

  test('appends the default /v1 prefix and the images endpoint', async () => {
    await generateImage(CONFIG, 'a circle');
    expect(global.fetch.mock.calls[0][0]).toBe('/openai/v1/images/generations');
  });

  test('strips trailing slashes from the base url', async () => {
    await generateImage({ ...CONFIG, url: '/openai/' }, 'a circle');
    expect(global.fetch.mock.calls[0][0]).toBe('/openai/v1/images/generations');
  });

  test('does not double-append a prefix the url already ends with', async () => {
    // Regression: a user pasting the conventional full endpoint got
    // /v1/v1/images/generations → 404 with no hint the URL was the problem.
    await generateImage({ ...CONFIG, url: 'https://api.example.com/v1' }, 'a circle');
    expect(global.fetch.mock.calls[0][0]).toBe('https://api.example.com/v1/images/generations');
  });

  test('honors a custom apiPath and does not double-append it either', async () => {
    await generateImage({ ...CONFIG, url: '/glm', apiPath: '/api/paas/v4' }, 'a circle');
    expect(global.fetch.mock.calls[0][0]).toBe('/glm/api/paas/v4/images/generations');

    await generateImage(
      { ...CONFIG, url: 'https://open.bigmodel.cn/api/paas/v4', apiPath: '/api/paas/v4' },
      'a circle',
    );
    expect(global.fetch.mock.calls[1][0]).toBe('https://open.bigmodel.cn/api/paas/v4/images/generations');
  });

  test('normalizes an apiPath given without a leading slash', async () => {
    await generateImage({ ...CONFIG, url: 'https://api.example.com', apiPath: 'v1' }, 'a circle');
    expect(global.fetch.mock.calls[0][0]).toBe('https://api.example.com/v1/images/generations');
  });
});

// ============================================================================
// Cancellation
// ============================================================================

describe('generateImage cancellation', () => {
  test('an already-aborted signal throws AbortError without calling fetch', async () => {
    const controller = new AbortController();
    controller.abort();

    const err = await generateImage(CONFIG, 'a circle', () => {}, controller.signal).catch((e) => e);

    expect(err.name).toBe('AbortError');
    expect(global.fetch).not.toHaveBeenCalled();
  });

  test('aborting mid-flight rejects and the listener is detached afterwards', async () => {
    const controller = new AbortController();
    global.fetch.mockImplementation((_url, opts) => new Promise((_resolve, reject) => {
      opts.signal.addEventListener('abort', () => {
        reject(new DOMException('The operation was aborted.', 'AbortError'));
      });
    }));

    const pending = generateImage(CONFIG, 'a circle', () => {}, controller.signal);
    controller.abort();

    const err = await pending.catch((e) => e);
    expect(err.name).toBe('AbortError');
    // A leaked listener would abort the next request through the same signal.
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  test('a timeout is reported as TimeoutError, distinct from cancellation', async () => {
    jest.useFakeTimers();
    global.fetch.mockImplementation((_url, opts) => new Promise((_resolve, reject) => {
      opts.signal.addEventListener('abort', () => {
        reject(new DOMException('The operation was aborted.', 'AbortError'));
      });
    }));

    const settled = generateImage(CONFIG, 'a circle', () => {}, undefined, 5000).catch((e) => e);
    jest.advanceTimersByTime(5001);
    const err = await settled;

    expect(err.name).toBe('TimeoutError');
    expect(err.message).toMatch(/timed out after 5s/);

    jest.useRealTimers();
  });
});

// ============================================================================
// Oversized payloads
// ============================================================================

describe('generateImage payload limit', () => {
  test('rejects a base64 payload beyond the insert cap with a size hint', async () => {
    const oversized = 'A'.repeat(8 * 1024 * 1024 + 1);
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: oversized }] }));

    const err = await generateImage(CONFIG, 'a circle').catch((e) => e);
    expect(err.message).toMatch(/too large to insert/);
    expect(err.message).toMatch(/Settings → Image Generation/);
  });
});

// ============================================================================
// Logging
// ============================================================================

describe('generateImage logging', () => {
  test('logs the model on start and the payload size on success', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'QUJD' }] }));
    const log = jest.fn();

    await generateImage(CONFIG, 'a circle', log);

    expect(log).toHaveBeenCalledWith('Generating image [gpt-image-1]...', 'info');
    expect(log).toHaveBeenCalledWith(expect.stringContaining('Image received'), 'success');
  });

  test('works without a log callback', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'QUJD' }] }));
    await expect(generateImage(CONFIG, 'a circle')).resolves.toHaveProperty('base64', 'QUJD');
  });
});

// ============================================================================
// testImageConnection
// ============================================================================

describe('testImageConnection', () => {
  test('reports ok with the model and payload size on success', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'A'.repeat(2048) }] }));

    const result = await testImageConnection(CONFIG);

    expect(result.ok).toBe(true);
    expect(result.detail).toContain('gpt-image-1');
    expect(result.detail).toContain('2 KB');
  });

  test('issues a real minimal generation rather than a ping', async () => {
    global.fetch.mockResolvedValue(jsonResponse({ data: [{ b64_json: 'QUJD' }] }));

    await testImageConnection(CONFIG);

    expect(global.fetch.mock.calls[0][0]).toBe('/openai/v1/images/generations');
    expect(bodyOf().prompt).toBeTruthy();
  });

  test('converts an HTTP failure into a report instead of throwing', async () => {
    global.fetch.mockResolvedValue({
      ok: false, status: 401, statusText: 'Unauthorized', text: async () => 'invalid api key',
    });

    const result = await testImageConnection(CONFIG);

    expect(result.ok).toBe(false);
    expect(result.detail).toContain('HTTP 401');
    expect(result.detail).toContain('invalid api key');
  });

  test('converts a configuration error into a report', async () => {
    const result = await testImageConnection({ url: '', model: '' });

    expect(result).toEqual({
      ok: false,
      detail: 'No image provider endpoint configured. Set one in Settings → Image Generation.',
    });
  });

  test('still propagates cancellation', async () => {
    const controller = new AbortController();
    controller.abort();

    const err = await testImageConnection(CONFIG, () => {}, controller.signal).catch((e) => e);

    // Not a {ok:false} report: the user cancelled, which is not a test result.
    expect(err.name).toBe('AbortError');
  });
});
