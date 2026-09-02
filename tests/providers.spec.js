/**
 * Tests for the provider catalog (src/lib/providers.js).
 * Pins the preset list, shapes, and OpenAI-prefix differences per provider.
 */

const {
  PROVIDER_PRESETS,
  KNOWN_PROVIDERS,
  getProviderPreset,
  defaultProviderConfig,
  isStaticHostOrigin,
} = require('../src/lib/providers.js');

describe('providers catalog', () => {
  test('exposes the expected provider ids in a stable order', () => {
    expect(KNOWN_PROVIDERS).toEqual(['ollama', 'vllm', 'deepseek', 'glm', 'kimi', 'minimax', 'minimax-cn', 'zhongkeyu', 'custom']);
  });

  test('every preset has label, url, apiPath, and model', () => {
    for (const id of KNOWN_PROVIDERS) {
      const preset = PROVIDER_PRESETS[id];
      expect(typeof preset.label).toBe('string');
      expect(preset.label.length).toBeGreaterThan(0);
      expect(typeof preset.apiPath).toBe('string');
      expect(preset.apiPath.startsWith('/')).toBe(true);
      if (id !== 'custom') {
        expect(preset.url.length).toBeGreaterThan(0);
        expect(preset.model.length).toBeGreaterThan(0);
      }
    }
  });

  test('local presets use the /v1 OpenAI prefix', () => {
    expect(PROVIDER_PRESETS.ollama.apiPath).toBe('/v1');
    expect(PROVIDER_PRESETS.vllm.apiPath).toBe('/v1');
  });

  test('cloud presets point at their documented API prefixes', () => {
    expect(PROVIDER_PRESETS.deepseek.apiPath).toBe('/v1');
    expect(PROVIDER_PRESETS.kimi.apiPath).toBe('/v1');
    expect(PROVIDER_PRESETS.minimax.apiPath).toBe('/v1');
    expect(PROVIDER_PRESETS['minimax-cn'].apiPath).toBe('/v1');
    expect(PROVIDER_PRESETS.zhongkeyu.apiPath).toBe('/v1');
    // Zhipu GLM serves OpenAI-compatible endpoints under /api/paas/v4
    expect(PROVIDER_PRESETS.glm.apiPath).toBe('/api/paas/v4');
  });

  test('cloud presets reference a key portal and local ones do not', () => {
    expect(PROVIDER_PRESETS.deepseek.keyHint).toContain('deepseek.com');
    expect(PROVIDER_PRESETS.glm.keyHint).toContain('bigmodel.cn');
    expect(PROVIDER_PRESETS.kimi.keyHint).toContain('moonshot.cn');
    expect(PROVIDER_PRESETS.minimax.keyHint).toContain('minimax.io');
    expect(PROVIDER_PRESETS['minimax-cn'].keyHint).toContain('minimax.cn');
    expect(PROVIDER_PRESETS.zhongkeyu.keyHint).toContain('zhongkeyu.com');
    expect(PROVIDER_PRESETS.ollama.keyHint).toBeUndefined();
    expect(PROVIDER_PRESETS.vllm.keyHint).toBeUndefined();
  });

  test('MiniMax exposes separate international and China presets', () => {
    // The two platforms have separate API origins and key portals.
    expect(PROVIDER_PRESETS.minimax.url).toBe('https://api.minimax.io');
    expect(PROVIDER_PRESETS['minimax-cn'].url).toBe('https://api.minimaxi.com');
    expect(PROVIDER_PRESETS.minimax.model).toBe('MiniMax-M3');
    expect(PROVIDER_PRESETS['minimax-cn'].model).toBe('MiniMax-M3');
    expect(PROVIDER_PRESETS['minimax-cn'].label).toContain('中国站');
  });

  test('zhongkeyu preset points at the gateway root with an OpenAI prefix', () => {
    // zhongkeyu.com is a New API gateway: OpenAI-compatible endpoints live
    // under /v1, keys are issued on the site itself.
    expect(PROVIDER_PRESETS.zhongkeyu.url).toBe('https://zhongkeyu.com');
    expect(PROVIDER_PRESETS.zhongkeyu.model.length).toBeGreaterThan(0);
    expect(PROVIDER_PRESETS.zhongkeyu.label).toContain('中科大模型');
  });

  test('cloud presets default to absolute HTTPS origins (statically hosted installs work serverless)', () => {
    // All six providers send Access-Control-Allow-Origin for the add-in's
    // hosted origins (verified), so the store install talks to them directly
    // instead of relying on same-origin proxy paths that only exist behind
    // the docker/dev server.
    expect(PROVIDER_PRESETS.deepseek.url).toBe('https://api.deepseek.com');
    expect(PROVIDER_PRESETS.glm.url).toBe('https://open.bigmodel.cn');
    expect(PROVIDER_PRESETS.kimi.url).toBe('https://api.moonshot.cn');
    expect(PROVIDER_PRESETS.minimax.url).toBe('https://api.minimax.io');
    expect(PROVIDER_PRESETS['minimax-cn'].url).toBe('https://api.minimaxi.com');
    expect(PROVIDER_PRESETS.zhongkeyu.url).toBe('https://zhongkeyu.com');
    for (const id of ['deepseek', 'glm', 'kimi', 'minimax', 'minimax-cn', 'zhongkeyu']) {
      expect(PROVIDER_PRESETS[id].staticOk).toBe(true);
    }
  });

  test('local-model presets default to same-origin proxy paths (static-host installs need a relay)', () => {
    // A static HTTPS page cannot reach http://localhost (mixed-content
    // blocking), so Ollama/vLLM keep the proxy-path default that resolves
    // when the add-in is served by the docker/dev server.
    expect(PROVIDER_PRESETS.ollama.url).toBe('/ollama');
    expect(PROVIDER_PRESETS.vllm.url).toBe('/vllm');
    expect(PROVIDER_PRESETS.ollama.staticOk).toBe(false);
    expect(PROVIDER_PRESETS.vllm.staticOk).toBe(false);
  });

  test('cloud presets default to absolute HTTPS origins on static hosts', () => {
    // All six providers return Access-Control-Allow-Origin for public
    // origins (verified), so the statically hosted install (marketplace /
    // GitHub Pages) calls them directly — no server of our own needed.
    const cfg = defaultProviderConfig('https://scilence2022.github.io');
    expect(cfg.deepseek.url).toBe('https://api.deepseek.com');
    expect(cfg.glm.url).toBe('https://open.bigmodel.cn');
    expect(cfg.kimi.url).toBe('https://api.moonshot.cn');
    expect(cfg.minimax.url).toBe('https://api.minimax.io');
    expect(cfg['minimax-cn'].url).toBe('https://api.minimaxi.com');
    expect(cfg.zhongkeyu.url).toBe('https://zhongkeyu.com');
  });

  test('local-served origins default cloud providers to same-origin proxy paths', () => {
    // The providers REFUSE CORS for localhost/private-IP origins (verified:
    // no Access-Control-Allow-Origin is emitted), so direct absolute calls
    // cannot work from a locally served taskpane — the proxy path is the
    // only mechanism there.
    for (const origin of ['', 'https://localhost:3001', 'https://192.168.1.63:3001']) {
      const cfg = defaultProviderConfig(origin);
      expect(cfg.deepseek.url).toBe('/deepseek');
      expect(cfg.glm.url).toBe('/glm');
      expect(cfg['minimax-cn'].url).toBe('/minimax-cn');
      expect(cfg.zhongkeyu.url).toBe('/zhongkeyu');
    }
  });

  test('local-model presets use proxy paths on every origin (mixed content blocks http)', () => {
    // WebKit (Word on Mac) blocks http://localhost from an HTTPS page with
    // no exemption (bugs.webkit.org 171934/173161), so there is no usable
    // absolute default for Ollama/vLLM.
    for (const origin of ['https://scilence2022.github.io', '']) {
      const cfg = defaultProviderConfig(origin);
      expect(cfg.ollama.url).toBe('/ollama');
      expect(cfg.vllm.url).toBe('/vllm');
    }
  });

  test('isStaticHostOrigin classifies hosted vs locally served origins', () => {
    expect(isStaticHostOrigin('https://scilence2022.github.io')).toBe(true);
    expect(isStaticHostOrigin('https://user.github.io')).toBe(true);
    expect(isStaticHostOrigin('https://localhost:3001')).toBe(false);
    expect(isStaticHostOrigin('https://192.168.1.63:3001')).toBe(false);
    expect(isStaticHostOrigin('')).toBe(false);
    expect(isStaticHostOrigin('not a url')).toBe(false);
  });

  test('getProviderPreset returns null for unknown ids', () => {
    expect(getProviderPreset('gpt4all')).toBeNull();
    expect(getProviderPreset('')).toBeNull();
  });

  test('defaultProviderConfig yields an editable entry per provider', () => {
    // Static origin: entry URLs equal the presets' absolute urls verbatim.
    const config = defaultProviderConfig('https://scilence2022.github.io');
    expect(Object.keys(config).sort()).toEqual([...KNOWN_PROVIDERS].sort());
    for (const id of KNOWN_PROVIDERS) {
      expect(config[id]).toEqual({
        url: PROVIDER_PRESETS[id].url,
        apiKey: '',
        model: PROVIDER_PRESETS[id].model,
        apiPath: PROVIDER_PRESETS[id].apiPath,
        thinkingLevel: 'default',
        temperature: 1,
      });
    }
  });

  test('defaultProviderConfig returns a fresh object per call', () => {
    const a = defaultProviderConfig();
    const b = defaultProviderConfig();
    expect(a).not.toBe(b);
    expect(a.ollama).not.toBe(b.ollama);
  });
});
