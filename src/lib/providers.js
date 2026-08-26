/**
 * Provider catalog for the LLM backend selector.
 *
 * Every provider speaks the OpenAI-compatible chat completions API, but the
 * base URL-to-endpoint mapping differs (Zhipu GLM uses /api/paas/v4, others
 * use /v1). The add-in talks to them through same-origin proxy paths
 * (/ollama, /deepseek, ...) served by the dev server or the production
 * server, which keeps the WebView free of cross-origin CORS pitfalls. A
 * custom provider lets users point at any OpenAI-compatible endpoint.
 *
 * @module providers
 */

/**
 * Provider presets keyed by backend id.
 *
 * Each entry:
 *   label    - UI display name
 *   url      - default base URL shown in Settings (proxy path or origin)
 *   apiPath  - API prefix appended to `url` for OpenAI endpoints
 *              ('' means the url already includes it)
 *   model    - default model id
 *   keyHint  - where to obtain an API key (optional, shown in the UI)
 *
 * Because webpack DefinePlugin replaces process.env.DEFAULT_* at build
 * time, Ollama/vLLM keep honoring the classic env overrides.
 *
 * @type {Object<string, {label: string, url: string, apiPath: string, model: string, keyHint?: string}>}
 */
export const PROVIDER_PRESETS = {
  ollama: {
    label: 'Ollama',
    url: process.env.DEFAULT_OLLAMA_URL || '/ollama',
    apiPath: '/v1',
    model: process.env.DEFAULT_MODEL || 'gpt-oss:20b',
  },
  vllm: {
    label: 'vLLM',
    url: process.env.DEFAULT_VLLM_URL || '/vllm',
    apiPath: '/v1',
    model: process.env.DEFAULT_VLLM_MODEL || 'qwen3.5-35b-a3b',
  },
  deepseek: {
    label: 'DeepSeek',
    url: '/deepseek',
    apiPath: '/v1',
    model: 'deepseek-chat',
    keyHint: 'platform.deepseek.com',
  },
  glm: {
    label: 'Zhipu GLM',
    url: '/glm',
    apiPath: '/api/paas/v4',
    model: 'glm-4.5',
    keyHint: 'open.bigmodel.cn',
  },
  kimi: {
    label: 'Moonshot Kimi',
    url: '/kimi',
    apiPath: '/v1',
    model: 'kimi-k2-0905',
    keyHint: 'platform.moonshot.cn',
  },
  // MiniMax runs two platforms with separate API origins and key portals:
  // international (api.minimax.io) and China (api.minimaxi.com).
  minimax: {
    label: 'MiniMax',
    url: '/minimax',
    apiPath: '/v1',
    model: 'MiniMax-M3',
    keyHint: 'platform.minimax.io',
  },
  'minimax-cn': {
    label: 'MiniMax 中国站',
    url: '/minimax-cn',
    apiPath: '/v1',
    model: 'MiniMax-M3',
    keyHint: 'platform.minimax.cn',
  },
  custom: {
    label: 'Custom (OpenAI-compatible)',
    url: '',
    apiPath: '/v1',
    model: '',
  },
};

/** @type {string[]} Ordered provider ids (mirrors the Settings select). */
export const KNOWN_PROVIDERS = Object.keys(PROVIDER_PRESETS);

/**
 * Returns the preset entry for a provider id (or null for unknown ids).
 *
 * @param {string} providerId
 * @returns {{label: string, url: string, apiPath: string, model: string, keyHint?: string}|null}
 */
export function getProviderPreset(providerId) {
  return PROVIDER_PRESETS[providerId] || null;
}

/**
 * Builds the default provider config object (used when no saved settings
 * exist and as the base for normalizeConfig).
 *
 * @returns {Object<string, {url: string, apiKey: string, model: string, apiPath: string}>}
 */
export function defaultProviderConfig() {
  /** @type {Object<string, {url: string, apiKey: string, model: string, apiPath: string}>} */
  const config = {};
  for (const id of KNOWN_PROVIDERS) {
    const preset = PROVIDER_PRESETS[id];
    config[id] = {
      url: preset.url,
      apiKey: '',
      model: preset.model,
      apiPath: preset.apiPath,
    };
  }
  return config;
}
