/**
 * Image Generation Client
 *
 * Network layer for text-to-image providers (see image-providers.js). Mirrors
 * llm-client.js in shape — same URL joining, same abort/timeout wiring, same
 * error-description style — but targets the images endpoint and returns raster
 * bytes instead of text.
 *
 * The result is raw raster base64, whichever way the provider answered:
 * Word's insertInlinePictureFromBase64 takes raster base64 only, so a
 * URL-returning provider has its bytes fetched here and a data-URL response is
 * unwrapped. The provider's raster format is preserved; the MIME-aware preview
 * and Word insertion path handle the bytes without exposing transport details.
 *
 * No Office.js. Uses fetch, so it is testable
 * under Jest with a fetch mock.
 *
 * @module image-client
 */

import { getImageProviderPreset, DEFAULT_IMAGE_SIZE } from './image-providers.js';

/** Default per-request timeout. Image models are slow — minutes, not seconds. */
const DEFAULT_IMAGE_TIMEOUT_MS = 180000;

/** Largest image payload accepted, in base64 characters (~6 MB binary). */
const MAX_IMAGE_B64_CHARS = 8 * 1024 * 1024;

/** MiniMax's documented maximum prompt length. */
const MINIMAX_MAX_PROMPT_CHARS = 1500;

/** Conservative cap for unknown OpenAI-compatible image relays. */
const GENERIC_MAX_PROMPT_CHARS = 4000;

const LOOPBACK_HOSTS = new Set(['localhost', '127.0.0.1', '[::1]']);
const ABSOLUTE_SCHEME_RE = /^[A-Za-z][A-Za-z0-9+.-]*:/;
const SENSITIVE_FIELD_RE = /((?:authorization|proxy-authorization|x-api-key|api[_-]?key|access[_-]?token|refresh[_-]?token|token|password|secret)\b\s*["']?\s*[:=]\s*)(?:(?:Bearer|Basic)\s+)?(?:"[^"]*"|'[^']*'|[^&\s,;}\]"']+)/gi;
const SENSITIVE_QUERY_RE = /([?&](?:token|access[_-]?token|refresh[_-]?token|api[_-]?key|x-api-key|authorization|password|secret|signature)=)[^&#\s"'<>]+/gi;

/**
 * Returns true for a same-origin relative path, any HTTPS URL, or a loopback
 * HTTP URL. Backslashes and protocol-relative references are rejected because
 * browsers can reinterpret them as cross-origin authorities.
 *
 * @param {*} value
 * @param {boolean} allowRelative
 * @returns {boolean}
 */
function _isAllowedImageUrl(value, allowRelative) {
    if (typeof value !== 'string') return false;
    const candidate = value.trim();
    // eslint-disable-next-line no-control-regex -- reject control characters before fetch
    if (!candidate || /[\u0000-\u001f\u007f]/.test(candidate)) return false;
    if (candidate.includes('\\') || candidate.startsWith('//')) return false;
    if (!ABSOLUTE_SCHEME_RE.test(candidate)) return allowRelative;
    if (!/^https?:\/\//i.test(candidate)) return false;

    let parsed;
    try {
        parsed = new URL(candidate);
    } catch (_err) {
        return false;
    }
    if (parsed.username || parsed.password) return false;
    if (parsed.protocol === 'https:') return true;
    if (parsed.protocol !== 'http:') return false;
    const hostname = parsed.hostname.toLowerCase().replace(/\.$/, '');
    return LOOPBACK_HOSTS.has(hostname);
}

/**
 * Validates an image-generation endpoint. Relative proxy paths are allowed;
 * absolute cleartext URLs are restricted to local loopback services.
 *
 * @param {*} value
 * @returns {boolean}
 */
export function isAllowedImageEndpoint(value) {
    return _isAllowedImageUrl(value, true);
}

/**
 * Validates a provider-hosted image URL. Only HTTPS and loopback HTTP are
 * accepted; callers handle approved base64 data URLs before this check.
 *
 * @param {*} value
 * @returns {boolean}
 */
export function isAllowedImageDownloadUrl(value) {
    return _isAllowedImageUrl(value, false);
}

/**
 * Redacts credentials, sensitive query/header values, and complete URLs from
 * provider-controlled error text. Keep useful status/provider wording while
 * preventing a backend from reflecting an API key or request URL.
 *
 * @param {*} message
 * @param {Array<*>} [secrets]
 * @returns {string}
 */
export function sanitizeImageErrorMessage(message, secrets = []) {
    let safe = String(message || '');
    for (const secretValue of secrets) {
        const secret = secretValue == null ? '' : String(secretValue);
        if (!secret) continue;
        safe = safe.split(secret).join('[redacted]');
        try {
            const encoded = encodeURIComponent(secret);
            if (encoded && encoded !== secret) safe = safe.split(encoded).join('[redacted]');
        } catch (_err) {
            // Literal replacement above still covers the configured secret.
        }
    }
    safe = safe.replace(SENSITIVE_FIELD_RE, '$1[redacted]');
    safe = safe.replace(SENSITIVE_QUERY_RE, '$1[redacted]');
    safe = safe.replace(/(Bearer\s+)[^\s,;)}\]"']+/gi, '$1[redacted]');
    safe = safe.replace(/\b(?:https?|ftp):\/\/[^\s"'<>]+/gi, '[redacted URL]');
    safe = safe.replace(/\bdata:[^\s"'<>]+/gi, '[redacted data URL]');
    return safe;
}

function _safeImageError(error, secrets = []) {
    const safe = new Error(
        sanitizeImageErrorMessage(error && error.message, secrets) || 'Image request failed.'
    );
    if (error && error.name) safe.name = sanitizeImageErrorMessage(error.name, secrets);
    return safe;
}

/**
 * Joins base URL, API prefix, and endpoint. Same contract as llm-client's
 * joinApiUrl: a base that already ends with the prefix is not suffixed twice,
 * so a user pasting `https://host/v1` does not get `/v1/v1/images/...`.
 *
 * @param {string} baseUrl
 * @param {string} apiPath
 * @param {string} endpoint
 * @returns {string}
 * @private
 */
function _joinApiUrl(baseUrl, apiPath, endpoint) {
    const base = String(baseUrl || '').replace(/\/+$/, '');
    const prefix = apiPath.startsWith('/') ? apiPath : `/${apiPath}`;
    const withPrefix = base.endsWith(prefix) ? base : base + prefix;
    return withPrefix + endpoint;
}

/**
 * Resolves the effective request shape for a config: explicit `apiFormat`
 * wins (custom relays, tests), otherwise the preset's.
 *
 * @param {object} config
 * @returns {string}
 * @private
 */
function _apiFormat(config) {
    if (config && config.apiFormat) return config.apiFormat;
    const preset = config && config.provider ? getImageProviderPreset(config.provider) : null;
    return (preset && preset.apiFormat) || 'openai-images';
}

/** Returns the selected provider preset, if any. */
function _preset(config) {
    return config && config.provider ? getImageProviderPreset(config.provider) : null;
}

/**
 * OpenAI's GPT Image models always return base64 and reject response_format;
 * DALL-E models and generic compatible relays use the older b64_json switch.
 */
function _shouldRequestOpenAiB64(config) {
    if (/^gpt-image-/i.test(String(config.model || ''))) return false;
    if (typeof config.wantsB64 === 'boolean') return config.wantsB64;
    const policy = config.responseFormat || (_preset(config) && _preset(config).responseFormat);
    return policy === 'dall-e-b64';
}

/** Maps the UI's pixel orientations to MiniMax aspect ratios. */
function _minimaxAspectRatio(size) {
    if (size === '1536x1024') return '3:2';
    if (size === '1024x1536') return '2:3';
    return '1:1';
}

/** Truncates at the provider's documented prompt boundary. */
function _boundedPrompt(config, prompt) {
    const format = _apiFormat(config);
    const limit = format === 'minimax-images'
        ? MINIMAX_MAX_PROMPT_CHARS
        : (/^gpt-image-/i.test(String(config.model || '')) ? 32000 : GENERIC_MAX_PROMPT_CHARS);
    if (prompt.length <= limit) return prompt;
    const suffix = '\n[Document context truncated]';
    return prompt.slice(0, Math.max(0, limit - suffix.length)) + suffix;
}

/**
 * Turns a failed response into a message carrying the provider's own
 * explanation (quota, content policy, bad model id — all things the user can
 * act on), truncated so a HTML error page cannot flood the activity log.
 *
 * @param {Response} response
 * @returns {Promise<string>}
 * @private
 */
async function _describeHttpError(response, secrets = []) {
    let detail = '';
    try {
        detail = sanitizeImageErrorMessage(await response.text(), secrets).slice(0, 300);
    } catch (_err) {
        // Body already consumed or unreadable — the status line still informs.
    }
    const statusText = sanitizeImageErrorMessage(response.statusText || '', secrets);
    const base = `Image request failed: HTTP ${response.status} ${statusText}`.trim();
    return detail ? `${base} — ${detail}` : base;
}

/**
 * Extracts raw raster base64 from whatever the provider returned.
 *
 * Handles the three observed shapes: `b64_json` (OpenAI with
 * response_format), a `url` pointing at the rendered file (CogView, and
 * OpenAI without the flag), and a `url` that is itself a data URL. URLs are
 * fetched here — the caller only ever sees raster bytes.
 *
 * @param {object} payload - Parsed JSON response body
 * @param {AbortSignal} [signal] - Propagated to the bytes fetch
 * @param {Array<*>} [secrets] - Values to remove from provider errors
 * @returns {Promise<string>} Raw raster base64, no data-URL prefix
 * @throws {Error} When the response carries no usable image
 * @private
 */
async function _extractImageBase64(payload, signal, secrets = []) {
    const minimaxStatus = payload && payload.base_resp;
    const minimaxStatusCode = minimaxStatus ? Number(minimaxStatus.status_code) : 0;
    if (minimaxStatus && minimaxStatusCode !== 0) {
        const detail = sanitizeImageErrorMessage(
            minimaxStatus.status_msg || 'unknown provider error',
            secrets,
        ).slice(0, 300);
        const safeStatusCode = Number.isFinite(minimaxStatusCode) ? minimaxStatusCode : 'unknown';
        throw new Error(`MiniMax image request failed (${safeStatusCode}): ${detail}`);
    }

    // MiniMax: { data: { image_base64: ["..."], image_urls: ["..."] } }
    if (payload && payload.data && !Array.isArray(payload.data)) {
        const inline = Array.isArray(payload.data.image_base64)
            ? payload.data.image_base64.find((value) => typeof value === 'string' && value)
            : '';
        if (inline) return inline;
        const hosted = Array.isArray(payload.data.image_urls)
            ? payload.data.image_urls.find((value) => typeof value === 'string' && value)
            : '';
        if (hosted) return _downloadImageBase64(hosted, signal);
    }

    const first = payload && Array.isArray(payload.data) ? payload.data[0] : null;
    if (!first) {
        throw new Error('The image provider returned no image data.');
    }

    if (typeof first.b64_json === 'string' && first.b64_json) {
        return first.b64_json;
    }

    const url = typeof first.url === 'string' ? first.url : '';
    if (!url) {
        throw new Error('The image provider returned neither base64 data nor a URL.');
    }

    // Data URL: unwrap without a network round trip.
    const dataUrlMatch = url.match(/^data:image\/[a-z0-9.+-]+;base64,(.*)$/i);
    if (dataUrlMatch) {
        return dataUrlMatch[1];
    }

    return _downloadImageBase64(url, signal);
}

/** Fetches a hosted image and converts its bytes to base64. */
async function _downloadImageBase64(url, signal) {
    const candidate = typeof url === 'string' ? url.trim() : '';

    // Data URL: unwrap without a network round trip. It is an inline provider
    // response, not a network destination, so it has its own narrow allowlist.
    const dataUrlMatch = candidate.match(/^data:image\/[a-z0-9.+-]+;base64,(.*)$/i);
    if (dataUrlMatch) return dataUrlMatch[1];

    if (!isAllowedImageDownloadUrl(candidate)) {
        throw new Error('The image provider returned an unsupported image URL.');
    }

    // A hosted file: fetch the bytes. This is a cross-origin GET against the
    // provider's CDN, which is why presets prefer inline base64 when available.
    let imgResponse;
    try {
        imgResponse = await fetch(candidate, { signal });
    } catch (err) {
        if (err && err.name === 'AbortError') throw err;
        throw new Error('The image provider image download failed.');
    }
    if (!imgResponse.ok) {
        throw new Error(
            `The image provider returned an image URL that could not be downloaded (HTTP ${imgResponse.status}). `
            + 'If this add-in is served over HTTPS from a static host, the provider CDN may be blocking the request.'
        );
    }
    let buffer;
    try {
        buffer = await imgResponse.arrayBuffer();
    } catch (err) {
        if (err && err.name === 'AbortError') throw err;
        throw new Error('The image provider image download failed.');
    }
    let binary = '';
    const bytes = new Uint8Array(buffer);
    // Chunked conversion: String.fromCharCode(...bytes) overflows the call
    // stack on megabyte payloads.
    const CHUNK = 0x8000;
    for (let i = 0; i < bytes.length; i += CHUNK) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
    }
    return btoa(binary);
}

/**
 * Builds the request body for one image generation.
 *
 * @param {object} config - { model, size, provider, apiFormat }
 * @param {string} prompt - The design brief
 * @returns {object}
 * @private
 */
function _buildBody(config, prompt) {
    const format = _apiFormat(config);
    const safePrompt = _boundedPrompt(config, prompt);
    if (format === 'minimax-images') {
        return {
            model: config.model,
            prompt: safePrompt,
            aspect_ratio: _minimaxAspectRatio(config.size),
            response_format: 'base64',
            n: 1,
            prompt_optimizer: true,
        };
    }

    const size = config.size && config.size !== 'auto' ? config.size : undefined;
    /** @type {Record<string, *>} */
    const body = {
        model: config.model,
        prompt: safePrompt,
        n: 1,
    };
    if (size) body.size = size;
    // CogView rejects response_format. GPT Image rejects it too because those
    // models always return base64; DALL-E/compatible relays need b64_json to
    // avoid a second, potentially CORS-blocked, CDN request.
    if (format === 'openai-images' && _shouldRequestOpenAiB64(config)) {
        body.response_format = 'b64_json';
    }
    return body;
}

/** Selects the endpoint path for the provider's wire format. */
function _imageEndpoint(config) {
    return _apiFormat(config) === 'minimax-images'
        ? '/image_generation'
        : '/images/generations';
}

/**
 * Generates ONE image from a text prompt and returns raw raster base64.
 *
 * Abort/timeout wiring mirrors llm-client (no AbortSignal.any — WebView2 lacks
 * it): a local controller drives the fetch, the external signal aborts it, and
 * a timeout abort is reported as TimeoutError so callers can tell the two
 * apart.
 *
 * @param {object} config - { url, apiKey, model, apiPath, size, provider, apiFormat? }
 * @param {string} prompt - Design brief for the image
 * @param {function} [log] - Logging callback (message, type)
 * @param {AbortSignal} [signal] - Cancellation signal
 * @param {number} [timeoutMs=180000] - Per-request timeout
 * @returns {Promise<{base64: string, model: string, size: string}>}
 * @throws {DOMException} AbortError on cancellation
 * @throws {Error} TimeoutError on timeout; Error with provider detail on HTTP failure
 */
export async function generateImage(config, prompt, log = () => {}, signal, timeoutMs = DEFAULT_IMAGE_TIMEOUT_MS) {
    if (!config || !config.url) {
        throw new Error('No image provider endpoint configured. Set one in Settings → Image Generation.');
    }
    if (!config.model) {
        throw new Error('No image model configured. Set one in Settings → Image Generation.');
    }
    if (typeof prompt !== 'string' || !prompt.trim()) {
        throw new Error('An image prompt is required.');
    }

    const endpoint = typeof config.url === 'string' ? config.url.trim() : '';
    if (!isAllowedImageEndpoint(endpoint)) {
        throw new Error('Image provider endpoint must be an HTTPS URL, a localhost HTTP URL, or a relative proxy path.');
    }
    const url = _joinApiUrl(endpoint, config.apiPath || '/v1', _imageEndpoint(config));
    const errorSecrets = [config.apiKey, url];
    const headers = { 'Content-Type': 'application/json' };
    if (config.apiKey) headers['Authorization'] = `Bearer ${config.apiKey}`;

    const body = JSON.stringify(_buildBody(config, prompt.trim()));

    const localController = new AbortController();
    let timedOut = false;
    const timeoutId = setTimeout(() => {
        timedOut = true;
        localController.abort();
    }, timeoutMs);

    let onExternalAbort;
    if (signal) {
        if (signal.aborted) {
            clearTimeout(timeoutId);
            throw new DOMException('The operation was aborted.', 'AbortError');
        }
        onExternalAbort = () => localController.abort();
        signal.addEventListener('abort', onExternalAbort);
    }

    try {
        log(`Generating image [${config.model}]...`, 'info');
        const response = await fetch(url, {
            method: 'POST',
            headers,
            body,
            signal: localController.signal,
        });
        if (!response.ok) {
            throw new Error(await _describeHttpError(response, errorSecrets));
        }
        const payload = await response.json();
        const base64 = await _extractImageBase64(payload, localController.signal, errorSecrets);
        if (base64.length > MAX_IMAGE_B64_CHARS) {
            throw new Error(
                `The generated image is too large to insert (${(base64.length / 1024 / 1024).toFixed(1)} MB). `
                + 'Try a smaller size in Settings → Image Generation.'
            );
        }
        log(`Image received (${(base64.length / 1024).toFixed(0)} KB base64).`, 'success');
        return {
            base64,
            model: config.model,
            size: config.size || DEFAULT_IMAGE_SIZE,
        };
    } catch (err) {
        if (timedOut && err.name === 'AbortError') {
            const timeoutErr = new Error(
                `Image generation timed out after ${Math.round(timeoutMs / 1000)}s`
            );
            timeoutErr.name = 'TimeoutError';
            throw timeoutErr;
        }
        if (err && err.name === 'AbortError') throw err;
        throw _safeImageError(err, errorSecrets);
    } finally {
        clearTimeout(timeoutId);
        if (signal && onExternalAbort) {
            signal.removeEventListener('abort', onExternalAbort);
        }
    }
}

/**
 * Probes an image provider config without generating a full image — the
 * Settings "Test" button. There is no cheap ping on these APIs, so this issues
 * a real minimal generation and reports whether bytes came back.
 *
 * @param {object} config
 * @param {function} [log]
 * @param {AbortSignal} [signal]
 * @returns {Promise<{ok: boolean, detail: string}>}
 */
export async function testImageConnection(config, log = () => {}, signal) {
    try {
        const result = await generateImage(config, 'a small grey circle on a white background', log, signal, 60000);
        return {
            ok: true,
            detail: `${config.model} responded with ${(result.base64.length / 1024).toFixed(0)} KB of image data.`,
        };
    } catch (err) {
        if (err.name === 'AbortError') throw err;
        return { ok: false, detail: err.message };
    }
}
