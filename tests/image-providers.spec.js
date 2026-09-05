/**
 * Image-provider catalog tests.
 *
 * The image catalog is a second, independent copy of the origin-adaptive
 * defaulting rule already covered for providers.js: a static install
 * (*.github.io) has no same-origin proxy, so cloud image APIs must be reached
 * at their absolute origins, while a locally served taskpane must go through
 * the proxy paths or CORS kills the request. Getting this backwards produces a
 * config that looks right in Settings and fails only at generation time, so
 * both directions are pinned here.
 *
 * Also covers the lookup helpers image-client.js depends on for request
 * shaping (getImageProviderPreset / imageSizesFor), including their
 * unknown-id fallbacks — a stale provider id in persisted config must not
 * throw.
 */

const {
  IMAGE_SIZES,
  DEFAULT_IMAGE_SIZE,
  IMAGE_PROVIDER_PRESETS,
  KNOWN_IMAGE_PROVIDERS,
  getImageProviderPreset,
  imageSizesFor,
  defaultImageProviderConfig,
} = require('../src/lib/image-providers.js');

const STATIC_ORIGIN = 'https://scilence2022.github.io';

describe('getImageProviderPreset', () => {
  test('returns the preset for each known id', () => {
    for (const id of KNOWN_IMAGE_PROVIDERS) {
      expect(getImageProviderPreset(id)).toBe(IMAGE_PROVIDER_PRESETS[id]);
    }
  });

  test('returns null for an unknown id', () => {
    expect(getImageProviderPreset('no-such-provider')).toBeNull();
  });

  test('returns null for nullish / empty ids', () => {
    expect(getImageProviderPreset('')).toBeNull();
    expect(getImageProviderPreset(undefined)).toBeNull();
    expect(getImageProviderPreset(null)).toBeNull();
  });
});

describe('imageSizesFor', () => {
  test('falls back to the shared list for an unknown id', () => {
    expect(imageSizesFor('no-such-provider')).toBe(IMAGE_SIZES);
  });

  test('falls back to the shared list for presets declaring no sizes', () => {
    // No current preset overrides `sizes`; the fallback is what Settings renders.
    expect(imageSizesFor('openai')).toBe(IMAGE_SIZES);
    expect(imageSizesFor('glm')).toBe(IMAGE_SIZES);
  });

  test('the shared list starts with auto and contains the default size', () => {
    expect(IMAGE_SIZES[0]).toBe('auto');
    expect(IMAGE_SIZES).toContain(DEFAULT_IMAGE_SIZE);
  });
});

describe('defaultImageProviderConfig origin adaptation', () => {
  test('a static host uses absolute API origins for proxy-backed providers', () => {
    const cfg = defaultImageProviderConfig(STATIC_ORIGIN);
    expect(cfg.glm.url).toBe('https://open.bigmodel.cn');
    expect(cfg.minimax.url).toBe('https://api.minimax.io');
    // OpenRouter and SiliconFlow also default to their absolute origins on
    // static hosts — both return Access-Control-Allow-Origin: * for public
    // origins (verified), so the marketplace install talks to them directly.
    expect(cfg.openrouter.url).toBe('https://openrouter.ai');
    expect(cfg.siliconflow.url).toBe('https://api.siliconflow.cn');
    // zhongkeyu.com mirrors its chat preset's staticOk flag.
    expect(cfg.zhongkeyu.url).toBe('https://zhongkeyu.com');
  });

  test('a locally served origin uses same-origin proxy paths', () => {
    for (const origin of ['', 'http://localhost:3000']) {
      const cfg = defaultImageProviderConfig(origin);
      expect(cfg.glm.url).toBe('/glm');
      expect(cfg.minimax.url).toBe('/minimax');
      expect(cfg.openrouter.url).toBe('/openrouter');
      expect(cfg.siliconflow.url).toBe('/siliconflow');
      expect(cfg.zhongkeyu.url).toBe('/zhongkeyu');
    }
  });

  test('providers without a proxyUrl keep their preset url on both hosts', () => {
    // OpenAI images has no absolute default: api.openai.com sends no CORS
    // headers, so the proxy path is the only value that can ever work.
    expect(defaultImageProviderConfig(STATIC_ORIGIN).openai.url).toBe('/openai');
    expect(defaultImageProviderConfig('').openai.url).toBe('/openai');
    expect(defaultImageProviderConfig(STATIC_ORIGIN).custom.url).toBe('');
  });

  test('yields one fully-formed entry per known provider', () => {
    const cfg = defaultImageProviderConfig('');
    expect(Object.keys(cfg)).toEqual(KNOWN_IMAGE_PROVIDERS);
    for (const id of KNOWN_IMAGE_PROVIDERS) {
      expect(cfg[id]).toEqual({
        url: expect.any(String),
        apiKey: '',
        model: IMAGE_PROVIDER_PRESETS[id].model,
        apiPath: IMAGE_PROVIDER_PRESETS[id].apiPath,
        size: DEFAULT_IMAGE_SIZE,
      });
    }
  });

  test('returns a fresh, mutable object per call', () => {
    const a = defaultImageProviderConfig('');
    const b = defaultImageProviderConfig('');
    expect(a).not.toBe(b);
    expect(a.glm).not.toBe(b.glm);
    a.glm.apiKey = 'typed-by-user';
    expect(b.glm.apiKey).toBe('');
  });

  test('an unparseable origin is treated as local-served', () => {
    // isStaticHostOrigin swallows the URL parse error; the safe default is
    // the proxy path, which works wherever our own server is present.
    expect(defaultImageProviderConfig('not a url').glm.url).toBe('/glm');
  });
});

describe('preset request-shape metadata', () => {
  test('every preset declares an apiFormat the client understands', () => {
    for (const id of KNOWN_IMAGE_PROVIDERS) {
      expect(['openai-images', 'glm-images', 'minimax-images'])
        .toContain(IMAGE_PROVIDER_PRESETS[id].apiFormat);
    }
  });

  test('OpenAI-compatible presets carry the DALL-E response policy', () => {
    // GPT Image models are handled as a special case by image-client.js, but
    // DALL-E and compatible relays need the b64_json switch to avoid a CDN GET.
    expect(IMAGE_PROVIDER_PRESETS.openai.responseFormat).toBe('dall-e-b64');
    expect(IMAGE_PROVIDER_PRESETS.custom.responseFormat).toBe('dall-e-b64');
  });

  test('CogView has no response_format policy', () => {
    // CogView rejects response_format outright; absence is intentional rather
    // than a legacy wantsB64=false flag.
    expect(IMAGE_PROVIDER_PRESETS.glm.responseFormat).toBeUndefined();
    expect(IMAGE_PROVIDER_PRESETS.glm.apiFormat).toBe('glm-images');
  });

  test('MiniMax declares its native endpoint format and aspect-ratio sizes', () => {
    expect(IMAGE_PROVIDER_PRESETS.minimax.apiFormat).toBe('minimax-images');
    expect(imageSizesFor('minimax')).toEqual(IMAGE_SIZES);
    expect(IMAGE_PROVIDER_PRESETS.minimax.model).toBe('image-01');
  });

  test('OpenRouter and SiliconFlow image presets reuse the OpenAI request shape', () => {
    // Both gateways expose /v1/images/generations (or /api/v1/images), so the
    // shared 'openai-images' format applies. Their upstream hosts answer
    // ACAO=* for public origins (verified), so static installs reach them
    // directly without a relay.
    expect(IMAGE_PROVIDER_PRESETS.openrouter.apiFormat).toBe('openai-images');
    expect(IMAGE_PROVIDER_PRESETS.openrouter.apiPath).toBe('/api/v1');
    expect(IMAGE_PROVIDER_PRESETS.openrouter.model.length).toBeGreaterThan(0);
    expect(IMAGE_PROVIDER_PRESETS.openrouter.responseFormat).toBe('dall-e-b64');

    expect(IMAGE_PROVIDER_PRESETS.siliconflow.apiFormat).toBe('openai-images');
    expect(IMAGE_PROVIDER_PRESETS.siliconflow.apiPath).toBe('/v1');
    expect(IMAGE_PROVIDER_PRESETS.siliconflow.model.length).toBeGreaterThan(0);
    expect(IMAGE_PROVIDER_PRESETS.siliconflow.responseFormat).toBe('dall-e-b64');
  });

  test('zhongkeyu rides the same gateway origin as its chat preset', () => {
    // zhongkeyu.com is a New API gateway: OpenAI-shaped /v1 surface, so the
    // shared 'openai-images' format applies and b64_json is requested for
    // DALL-E-compatible relays (gpt-image-* is special-cased in the client).
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.apiFormat).toBe('openai-images');
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.apiPath).toBe('/v1');
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.url).toBe('https://zhongkeyu.com');
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.proxyUrl).toBe('/zhongkeyu');
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.staticOk).toBe(true);
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.model.length).toBeGreaterThan(0);
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.responseFormat).toBe('dall-e-b64');
    expect(IMAGE_PROVIDER_PRESETS.zhongkeyu.label).toContain('中科云');
  });
});
