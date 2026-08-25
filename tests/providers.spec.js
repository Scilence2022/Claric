/**
 * Tests for the provider catalog (src/lib/providers.js).
 * Pins the preset list, shapes, and OpenAI-prefix differences per provider.
 */

const {
  PROVIDER_PRESETS,
  KNOWN_PROVIDERS,
  getProviderPreset,
  defaultProviderConfig,
} = require('../src/lib/providers.js');

describe('providers catalog', () => {
  test('exposes the expected provider ids in a stable order', () => {
    expect(KNOWN_PROVIDERS).toEqual(['ollama', 'vllm', 'deepseek', 'glm', 'kimi', 'minimax', 'minimax-cn', 'custom']);
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
    // Zhipu GLM serves OpenAI-compatible endpoints under /api/paas/v4
    expect(PROVIDER_PRESETS.glm.apiPath).toBe('/api/paas/v4');
  });

  test('cloud presets reference a key portal and local ones do not', () => {
    expect(PROVIDER_PRESETS.deepseek.keyHint).toContain('deepseek.com');
    expect(PROVIDER_PRESETS.glm.keyHint).toContain('bigmodel.cn');
    expect(PROVIDER_PRESETS.kimi.keyHint).toContain('moonshot.cn');
    expect(PROVIDER_PRESETS.minimax.keyHint).toContain('minimax.io');
    expect(PROVIDER_PRESETS['minimax-cn'].keyHint).toContain('minimax.cn');
    expect(PROVIDER_PRESETS.ollama.keyHint).toBeUndefined();
    expect(PROVIDER_PRESETS.vllm.keyHint).toBeUndefined();
  });

  test('MiniMax exposes separate international and China presets', () => {
    // The two platforms have separate API origins and key portals.
    expect(PROVIDER_PRESETS.minimax.url).toBe('/minimax');
    expect(PROVIDER_PRESETS['minimax-cn'].url).toBe('/minimax-cn');
    expect(PROVIDER_PRESETS.minimax.model).toBe('MiniMax-M3');
    expect(PROVIDER_PRESETS['minimax-cn'].model).toBe('MiniMax-M3');
    expect(PROVIDER_PRESETS['minimax-cn'].label).toContain('中国站');
  });

  test('getProviderPreset returns null for unknown ids', () => {
    expect(getProviderPreset('gpt4all')).toBeNull();
    expect(getProviderPreset('')).toBeNull();
  });

  test('defaultProviderConfig yields an editable entry per provider', () => {
    const config = defaultProviderConfig();
    expect(Object.keys(config).sort()).toEqual([...KNOWN_PROVIDERS].sort());
    for (const id of KNOWN_PROVIDERS) {
      expect(config[id]).toEqual({
        url: PROVIDER_PRESETS[id].url,
        apiKey: '',
        model: PROVIDER_PRESETS[id].model,
        apiPath: PROVIDER_PRESETS[id].apiPath,
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
