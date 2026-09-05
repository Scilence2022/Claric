/**
 * Image-generation provider catalog.
 *
 * Separate from providers.js (the chat/LLM catalog) because the two are
 * independent choices: a user may run a local Ollama chat model while
 * generating artwork through a cloud image API, and most chat providers
 * expose no image endpoint at all. Settings therefore keeps its own image
 * provider + model + key, and the illustration pipeline reads this catalog.
 *
 * Every preset declares its request shape through `apiFormat`:
 *   'openai-images'  POST {base}{apiPath}/images/generations. GPT Image
 *                    models always return base64 and reject response_format;
 *                    DALL-E and compatible relays may request b64_json.
 *   'glm-images'     Zhipu's variant of the same endpoint. It accepts no
 *                    response_format and answers with a hosted URL.
 *   'minimax-images' POST {base}{apiPath}/image_generation with MiniMax's
 *                    aspect_ratio / response_format='base64' contract and
 *                    data.image_base64[] response shape.
 *
 * Origin-adaptive defaults follow the same rule as providers.js: same-origin
 * proxy paths when the add-in is served by its own server (cloud image APIs
 * refuse CORS for localhost origins exactly like their chat counterparts),
 * absolute API origins on a static host. See providers.js for the full
 * reasoning and the verified CORS behavior.
 *
 * @module image-providers
 */

import { isStaticHostOrigin } from './providers.js';

/**
 * The origin the taskpane is served from ('' outside a browser). Duplicated
 * from providers.js rather than exported across modules: it is three lines,
 * and importing it would couple the two catalogs for no benefit.
 *
 * @returns {string}
 * @private
 */
function runtimeOrigin() {
    try {
        return (typeof window !== 'undefined' && window.location && window.location.origin) || '';
    } catch (_err) {
        return '';
    }
}

/**
 * Image sizes offered in Settings. Kept to a small, widely-supported set:
 * every listed provider accepts these three, and a document illustration
 * rarely needs finer control. 'auto' lets the provider pick (gpt-image-1
 * supports it; others fall back to their own default).
 *
 * @type {readonly string[]}
 */
export const IMAGE_SIZES = Object.freeze(['auto', '1024x1024', '1536x1024', '1024x1536']);

/** Default size when config carries nothing usable. */
export const DEFAULT_IMAGE_SIZE = '1024x1024';

/**
 * Image provider presets keyed by id. Field meanings match providers.js
 * (`url` / `proxyUrl` / `apiPath` / `model` / `keyHint` / `staticOk`), plus:
 *
 *   apiFormat  - request/response shape ('openai-images' | 'glm-images' |
 *                'minimax-images')
 *   responseFormat - optional response-format policy. 'dall-e-b64' means
 *                request b64_json only for DALL-E models; GPT Image always
 *                returns base64 and rejects that field.
 *   sizes      - provider-specific size list; omitted means IMAGE_SIZES
 *
 * @type {Object<string, {label: string, url: string, proxyUrl?: string,
 *   apiPath: string, model: string, keyHint?: string, staticOk?: boolean,
 *   apiFormat: string, responseFormat?: string, sizes?: readonly string[]}>}
 */
export const IMAGE_PROVIDER_PRESETS = {
    openai: {
        label: 'OpenAI Images',
        // api.openai.com sends no CORS headers, so the proxy path is the only
        // default that can work from a locally served page (same as the chat
        // preset in providers.js).
        url: '/openai',
        apiPath: '/v1',
        model: 'gpt-image-1',
        keyHint: 'platform.openai.com',
        staticOk: false,
        apiFormat: 'openai-images',
        responseFormat: 'dall-e-b64',
    },
    glm: {
        label: 'Zhipu CogView',
        url: 'https://open.bigmodel.cn',
        proxyUrl: '/glm',
        apiPath: '/api/paas/v4',
        model: 'cogview-4',
        keyHint: 'open.bigmodel.cn',
        staticOk: true,
        apiFormat: 'glm-images',
        // CogView ignores response_format and always returns a URL.
    },
    minimax: {
        label: 'MiniMax Images',
        url: 'https://api.minimax.io',
        proxyUrl: '/minimax',
        apiPath: '/v1',
        model: 'image-01',
        keyHint: 'platform.minimax.io',
        staticOk: true,
        apiFormat: 'minimax-images',
        // MiniMax accepts aspect ratios rather than arbitrary pixel sizes.
        // Keep the UI's three document-friendly orientations plus auto.
        sizes: IMAGE_SIZES,
    },
    // zhongkeyu.com ("中科云") is the same New API gateway as the chat preset
    // in providers.js: one OpenAI-compatible /v1 surface relaying many
    // upstreams, so image generation rides /v1/images/generations. New API
    // relays commonly expose gpt-image-* (the client omits response_format
    // for those) plus DALL-E-compatible models (b64_json requested); the
    // model field stays user-editable for whatever the gateway account
    // enables. The staticOk flag mirrors the chat preset.
    zhongkeyu: {
        label: '中科云 Images',
        url: 'https://zhongkeyu.com',
        proxyUrl: '/zhongkeyu',
        apiPath: '/v1',
        model: 'gpt-image-1',
        keyHint: 'zhongkeyu.com',
        staticOk: true,
        apiFormat: 'openai-images',
        responseFormat: 'dall-e-b64',
    },
    // OpenRouter routes text-to-image requests through /api/v1/images (a
    // variant of the OpenAI /v1/images/generations shape — see the API
    // reference). Their endpoint returns b64_json directly, so the standard
    // OpenAI request body (with response_format=b64_json) works for models
    // like google/gemini-2.5-flash-image-preview, bytedance-seed/seedream,
    // black-forest-labs/flux*, and recraft/recraft-v4. openrouter.ai returns
    // Access-Control-Allow-Origin: * for public origins (verified).
    openrouter: {
        label: 'OpenRouter Images',
        url: 'https://openrouter.ai',
        proxyUrl: '/openrouter',
        apiPath: '/api/v1',
        model: 'google/gemini-2.5-flash-image-preview',
        keyHint: 'openrouter.ai',
        staticOk: true,
        apiFormat: 'openai-images',
        // OpenRouter accepts and prefers b64_json for downstream providers
        // that support it; some still return URLs, which the client handles.
        responseFormat: 'dall-e-b64',
    },
    // SiliconFlow's image API is OpenAI-shaped: POST /v1/images/generations
    // with model/prompt/size/n. Default model Kwai-Kolors/Kolors accepts
    // standard pixel sizes and returns a hosted URL (the OpenAI relay shape);
    // users can switch to Qwen-Image, stable-diffusion-3, etc. without
    // changing the provider. api.siliconflow.cn returns ACAO=* (verified).
    siliconflow: {
        label: 'SiliconFlow Images',
        url: 'https://api.siliconflow.cn',
        proxyUrl: '/siliconflow',
        apiPath: '/v1',
        model: 'Kwai-Kolors/Kolors',
        keyHint: 'cloud.siliconflow.cn',
        staticOk: true,
        apiFormat: 'openai-images',
        // SiliconFlow's hosted URL response matches the standard OpenAI
        // relay; we ask for b64_json when the model supports it (most do)
        // to avoid a second CDN GET that the local-server proxy cannot relay.
        responseFormat: 'dall-e-b64',
    },
    custom: {
        label: 'Custom (OpenAI-compatible)',
        url: '',
        apiPath: '/v1',
        model: '',
        staticOk: true,
        apiFormat: 'openai-images',
        // Custom OpenAI-compatible relays commonly implement the DALL-E
        // response_format contract. Users pointing this at a GPT Image relay
        // can name a gpt-image-* model and the client omits the field.
        responseFormat: 'dall-e-b64',
    },
};

/** @type {string[]} Ordered image provider ids (mirrors the Settings select). */
export const KNOWN_IMAGE_PROVIDERS = Object.keys(IMAGE_PROVIDER_PRESETS);

/**
 * Returns the preset for an image provider id, or null for unknown ids.
 *
 * @param {string} providerId
 * @returns {object|null}
 */
export function getImageProviderPreset(providerId) {
    return Object.prototype.hasOwnProperty.call(IMAGE_PROVIDER_PRESETS, providerId)
        ? IMAGE_PROVIDER_PRESETS[providerId]
        : null;
}

/**
 * Sizes a given provider accepts.
 *
 * @param {string} providerId
 * @returns {readonly string[]}
 */
export function imageSizesFor(providerId) {
    const preset = getImageProviderPreset(providerId);
    return (preset && preset.sizes) || IMAGE_SIZES;
}

/**
 * Builds the default image-provider config map, origin-adaptive like
 * defaultProviderConfig() in providers.js.
 *
 * @param {string} [origin] - Serving origin; defaults to the runtime origin
 *   ('' outside a browser = local-served, proxy defaults)
 * @returns {Object<string, {url: string, apiKey: string, model: string,
 *   apiPath: string, size: string}>}
 */
export function defaultImageProviderConfig(origin = runtimeOrigin()) {
    const staticHost = isStaticHostOrigin(origin);
    /** @type {Object<string, {url: string, apiKey: string, model: string, apiPath: string, size: string}>} */
    const config = {};
    for (const id of KNOWN_IMAGE_PROVIDERS) {
        const preset = IMAGE_PROVIDER_PRESETS[id];
        config[id] = {
            url: !staticHost && preset.proxyUrl ? preset.proxyUrl : preset.url,
            apiKey: '',
            model: preset.model,
            apiPath: preset.apiPath,
            size: DEFAULT_IMAGE_SIZE,
        };
    }
    return config;
}
