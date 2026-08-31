/**
 * Provider catalog for the LLM backend selector.
 *
 * Every provider speaks the OpenAI-compatible chat completions API, but the
 * base URL-to-endpoint mapping differs (Zhipu GLM uses /api/paas/v4, others
 * use /v1).
 *
 * Cloud providers default to their ABSOLUTE API origins: every one of them
 * returns Access-Control-Allow-Origin for the add-in's hosted origins
 * (verified against the Microsoft Marketplace / GitHub Pages origin), so a
 * statically hosted add-in (marketplace install) works with zero server of
 * its own. The same-origin proxy paths (/deepseek, /glm, ...) remain valid
 * endpoints for users running the docker/dev server — the endpoint field is
 * editable, and the docker server's proxy exists mainly for LOCAL model
 * hosts (Ollama/vLLM), which a static HTTPS page cannot reach directly
 * (mixed-content blocking).
 *
 * A custom provider lets users point at any OpenAI-compatible endpoint.
 *
 * @module providers
 */

/**
 * Provider presets keyed by backend id.
 *
 * Each entry:
 *   label    - UI display name
 *   url      - default base URL shown in Settings (absolute origin or
 *              same-origin proxy path)
 *   apiPath  - API prefix appended to `url` for OpenAI endpoints
 *              ('' means the url already includes it)
 *   model    - default model id
 *   keyHint  - where to obtain an API key (optional, shown in the UI)
 *   staticOk - true when the default works from a statically hosted origin
 *              (absolute URL, provider sends CORS headers). false marks
 *              local-model presets whose same-origin proxy default only
 *              resolves when the add-in is served by the docker/dev server.
 *
 * Because webpack DefinePlugin replaces process.env.DEFAULT_* at build
 * time, Ollama/vLLM keep honoring the classic env overrides.
 *
 * @type {Object<string, {label: string, url: string, apiPath: string, model: string, keyHint?: string, staticOk?: boolean}>}
 */
export const PROVIDER_PRESETS = {
  ollama: {
    label: 'Ollama',
    url: process.env.DEFAULT_OLLAMA_URL || '/ollama',
    apiPath: '/v1',
    model: process.env.DEFAULT_MODEL || 'gpt-oss:20b',
    staticOk: false,
  },
  vllm: {
    label: 'vLLM',
    url: process.env.DEFAULT_VLLM_URL || '/vllm',
    apiPath: '/v1',
    model: process.env.DEFAULT_VLLM_MODEL || 'qwen3.5-35b-a3b',
    staticOk: false,
  },
  deepseek: {
    label: 'DeepSeek',
    url: 'https://api.deepseek.com',
    apiPath: '/v1',
    model: 'deepseek-chat',
    keyHint: 'platform.deepseek.com',
    staticOk: true,
  },
  glm: {
    label: 'Zhipu GLM',
    url: 'https://open.bigmodel.cn',
    apiPath: '/api/paas/v4',
    model: 'glm-4.5',
    keyHint: 'open.bigmodel.cn',
    staticOk: true,
  },
  kimi: {
    label: 'Moonshot Kimi',
    url: 'https://api.moonshot.cn',
    apiPath: '/v1',
    model: 'kimi-k2-0905',
    keyHint: 'platform.moonshot.cn',
    staticOk: true,
  },
  // MiniMax runs two platforms with separate API origins and key portals:
  // international (api.minimax.io) and China (api.minimaxi.com).
  minimax: {
    label: 'MiniMax',
    url: 'https://api.minimax.io',
    apiPath: '/v1',
    model: 'MiniMax-M3',
    keyHint: 'platform.minimax.io',
    staticOk: true,
  },
  'minimax-cn': {
    label: 'MiniMax 中国站',
    url: 'https://api.minimaxi.com',
    apiPath: '/v1',
    model: 'MiniMax-M3',
    keyHint: 'platform.minimax.cn',
    staticOk: true,
  },
  custom: {
    label: 'Custom (OpenAI-compatible)',
    url: '',
    apiPath: '/v1',
    model: '',
    staticOk: true,
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
