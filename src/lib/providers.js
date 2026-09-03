/**
 * Provider catalog for the LLM backend selector.
 *
 * Every provider speaks the OpenAI-compatible chat completions API, but the
 * base URL-to-endpoint mapping differs (Zhipu GLM uses /api/paas/v4, others
 * use /v1).
 *
 * Defaults are ORIGIN-ADAPTIVE (see defaultProviderConfig):
 *  - On a static host (GitHub Pages / marketplace install) cloud providers
 *    default to their absolute API origins and are called directly — all
 *    six return CORS headers for public origins (verified). No server of
 *    our own is needed.
 *  - When the add-in is served by the local dev/production server, cloud
 *    providers default to same-origin proxy paths (/deepseek, /glm, ...),
 *    because those same providers REFUSE CORS for localhost/private-IP
 *    origins (verified) — direct calls cannot work from a locally served
 *    page.
 *  - Local models (Ollama/vLLM) always default to proxy paths: an HTTPS
 *    page cannot call http://localhost (mixed content — WebKit blocks it
 *    with no exemption, bugs.webkit.org 171934/173161), so direct local
 *    calls are not a portable option. A user who prefers direct calls can
 *    front the backend with HTTPS (OLLAMA_ORIGINS for CORS) and enter the
 *    absolute URL in Settings.
 *
 * A custom provider lets users point at any OpenAI-compatible endpoint.
 *
 * @module providers
 */

/** Hostnames matching this pattern serve the add-in without our backend. */
const STATIC_HOST_RE = /(^|\.)github\.io$/i;

/**
 * True when `origin` is a static hosting origin (no same-origin proxy
 * available). Pure function — exported for tests.
 *
 * @param {string} origin - An origin like 'https://scilence2022.github.io'
 * @returns {boolean}
 */
export function isStaticHostOrigin(origin) {
    if (!origin) return false;
    try {
        return STATIC_HOST_RE.test(new URL(origin).hostname);
    } catch (_err) {
        return false;
    }
}

/**
 * The origin the taskpane is being served from ('' outside a browser —
 * tests, module consumers without a DOM). Callers treat '' as
 * local-served, which keeps the historical proxy-path defaults.
 *
 * @returns {string}
 */
function runtimeOrigin() {
    try {
        return (typeof window !== 'undefined' && window.location && window.location.origin) || '';
    } catch (_err) {
        return '';
    }
}

/**
 * Provider presets keyed by backend id.
 *
 * Each entry:
 *   label    - UI display name
 *   url      - default base URL on a STATIC host (absolute API origin for
 *              cloud providers; proxy path for local models, which have no
 *              usable absolute default)
 *   proxyUrl - default base URL when served by the local server (same-origin
 *              proxy path). Cloud presets only; omitted when `url` applies
 *              everywhere.
 *   apiPath  - API prefix appended to `url`/`proxyUrl` for OpenAI endpoints
 *              ('' means the url already includes it)
 *   model    - default model id
 *   keyHint  - where to obtain an API key (optional, shown in the UI)
 *   staticOk - true when the preset is usable from a statically hosted
 *              install; false marks providers the browser cannot call
 *              directly (local models: mixed content; OpenAI: no CORS)
 *   staticHint - optional endpoint-hint override for staticOk:false
 *              providers whose restriction is not the local-model one
 *   apiFormat - 'anthropic' switches the LLM client to the Claude Messages
 *              API; omitted means OpenAI-compatible chat completions
 *
 * Because webpack DefinePlugin replaces process.env.DEFAULT_* at build
 * time, Ollama/vLLM keep honoring the classic env overrides.
 *
 * @type {Object<string, {label: string, url: string, proxyUrl?: string, apiPath: string, model: string, keyHint?: string, staticOk?: boolean, staticHint?: string, apiFormat?: string}>}
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
  openai: {
    label: 'OpenAI',
    // api.openai.com emits no Access-Control-Allow-Origin, so direct browser
    // calls fail on every origin — like the local-model presets, the default
    // is the same-origin proxy path on every install, and the hint explains
    // that a CORS-enabled relay URL is the alternative.
    url: '/openai',
    apiPath: '/v1',
    model: 'gpt-5.1',
    keyHint: 'platform.openai.com',
    staticOk: false,
    staticHint: 'OpenAI base URL — the default proxy path works when this add-in is served by its local server (api.openai.com allows no direct browser calls: it sends no CORS headers). From a static install (e.g. marketplace), enter an HTTPS OpenAI-compatible relay URL.',
  },
  claude: {
    label: 'Claude (Anthropic)',
    url: 'https://api.anthropic.com',
    proxyUrl: '/claude',
    apiPath: '/v1',
    model: 'claude-sonnet-4-6',
    keyHint: 'console.anthropic.com',
    // The client sends anthropic-dangerous-direct-browser-access, which opts
    // the API into browser CORS, so static installs can call it directly.
    staticOk: true,
    apiFormat: 'anthropic',
  },
  deepseek: {
    label: 'DeepSeek',
    url: 'https://api.deepseek.com',
    proxyUrl: '/deepseek',
    apiPath: '/v1',
    model: 'deepseek-chat',
    keyHint: 'platform.deepseek.com',
    staticOk: true,
  },
  glm: {
    label: 'Zhipu GLM',
    url: 'https://open.bigmodel.cn',
    proxyUrl: '/glm',
    apiPath: '/api/paas/v4',
    model: 'glm-4.5',
    keyHint: 'open.bigmodel.cn',
    staticOk: true,
  },
  kimi: {
    label: 'Moonshot Kimi',
    url: 'https://api.moonshot.cn',
    proxyUrl: '/kimi',
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
    proxyUrl: '/minimax',
    apiPath: '/v1',
    model: 'MiniMax-M3',
    keyHint: 'platform.minimax.io',
    staticOk: true,
  },
  'minimax-cn': {
    label: 'MiniMax 中国站',
    url: 'https://api.minimaxi.com',
    proxyUrl: '/minimax-cn',
    apiPath: '/v1',
    model: 'MiniMax-M3',
    keyHint: 'platform.minimax.cn',
    staticOk: true,
  },
  // zhongkeyu.com ("中科大模型-企业版") is a New API gateway that relays
  // many upstream models (GLM, DeepSeek, Claude, Gemini, GPT, ...) behind
  // one OpenAI-compatible /v1 surface; the default model is its cheapest
  // non-preview tier.
  zhongkeyu: {
    label: '中科大模型',
    url: 'https://zhongkeyu.com',
    proxyUrl: '/zhongkeyu',
    apiPath: '/v1',
    model: 'glm-5.3-flash',
    keyHint: 'zhongkeyu.com',
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
 * @returns {{label: string, url: string, apiPath: string, model: string, keyHint?: string, proxyUrl?: string, staticOk?: boolean, staticHint?: string, apiFormat?: string}|null}
 */
export function getProviderPreset(providerId) {
  return Object.prototype.hasOwnProperty.call(PROVIDER_PRESETS, providerId)
    ? PROVIDER_PRESETS[providerId]
    : null;
}

/**
 * Builds the default provider config object (used when no saved settings
 * exist and as the base for normalizeConfig).
 *
 * Cloud provider defaults are origin-adaptive: absolute API origins on
 * static hosts (direct CORS calls work there — providers reflect public
 * origins), same-origin proxy paths when served by the local server
 * (providers refuse CORS for localhost/private-IP origins, so direct calls
 * cannot work there). Local-model presets always use proxy paths.
 *
 * @param {string} [origin] - The serving origin; defaults to the runtime
 *   window origin ('' outside a browser = local-served, proxy defaults)
 * @returns {Object<string, {url: string, apiKey: string, model: string, apiPath: string, thinkingLevel: string, temperature: number}>}
 */
export function defaultProviderConfig(origin = runtimeOrigin()) {
    const staticHost = isStaticHostOrigin(origin);
    /** @type {Object<string, {url: string, apiKey: string, model: string, apiPath: string, thinkingLevel: string, temperature: number}>} */
    const config = {};
    for (const id of KNOWN_PROVIDERS) {
        const preset = PROVIDER_PRESETS[id];
        config[id] = {
            url: !staticHost && preset.proxyUrl ? preset.proxyUrl : preset.url,
            apiKey: '',
            model: preset.model,
            apiPath: preset.apiPath,
            thinkingLevel: 'default',
            temperature: 1,
        };
    }
    return config;
}
