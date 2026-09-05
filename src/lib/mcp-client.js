/**
 * MCP Client (Streamable HTTP)
 *
 * Minimal Model Context Protocol client for the browser/WebView2 taskpane:
 * JSON-RPC over the Streamable HTTP transport (2025-06-18 protocol).
 * Browser sandboxes have no stdio, so Streamable HTTP is the only transport
 * this add-in can speak; CORS/mixed-content are handled by the same-origin
 * proxy infrastructure (see scripts/docker-server.cjs, MCP_PROXY_PATH).
 *
 * The transport keeps one JSON-RPC request per POST and reads the reply
 * from either a plain JSON body or an SSE-encoded body (both allowed by
 * the spec); for SSE the first message matching the request id is taken.
 * Subscriptions/progress streams are out of scope here.
 *
 * Pure module — fetch is injectable, so the client is hermetic-testable.
 *
 * @module mcp-client
 */

/** Protocol version this client speaks. */
export const MCP_PROTOCOL_VERSION = '2025-06-18';

const LOOPBACK_HOSTS = new Set(['localhost', '127.0.0.1', '[::1]']);
const ABSOLUTE_SCHEME_RE = /^[A-Za-z][A-Za-z0-9+.-]*:/;

/**
 * Returns true for an HTTPS URL, a loopback HTTP URL, or a same-origin
 * relative reference. Relative references are intentionally not resolved:
 * browser fetch resolves them against the taskpane origin.
 *
 * @param {*} value
 * @returns {boolean}
 */
export function isAllowedMcpUrl(value) {
    if (typeof value !== 'string') return false;
    const candidate = value.trim();
    // eslint-disable-next-line no-control-regex -- reject control characters before fetch
    if (!candidate || /[\u0000-\u001f\u007f]/.test(candidate)) return false;
    // WHATWG URL parsing treats backslashes as authority separators for HTTP.
    // Reject them everywhere so a relative path cannot become cross-origin.
    if (candidate.includes('\\') || candidate.startsWith('//')) return false;
    if (!ABSOLUTE_SCHEME_RE.test(candidate)) return true;
    if (!/^https?:\/\//i.test(candidate)) return false;

    let parsed;
    try {
        parsed = new URL(candidate);
    } catch (_err) {
        return false;
    }
    if (!parsed.hostname || parsed.username || parsed.password) return false;
    if (parsed.protocol === 'https:') return true;
    if (parsed.protocol !== 'http:') return false;
    const hostname = parsed.hostname.toLowerCase().replace(/\.$/, '');
    return LOOPBACK_HOSTS.has(hostname);
}

/**
 * Removes credentials and sensitive query/header values from an error string.
 * The caller still gets a useful status/code, but a token cannot be echoed by
 * a browser fetch implementation or an upstream error response.
 *
 * @param {*} message
 * @param {*} [token]
 * @returns {string}
 */
export function sanitizeMcpErrorMessage(message, token) {
    let safe = String(message || 'MCP request failed');
    const secret = token == null ? '' : String(token);
    if (secret) {
        safe = safe.split(secret).join('[redacted]');
        try {
            const encoded = encodeURIComponent(secret);
            if (encoded && encoded !== secret) safe = safe.split(encoded).join('[redacted]');
        } catch (_err) {
            // A malformed value is still covered by the literal replacement.
        }
    }
    safe = safe.replace(/((?:authorization|x-api-key|api[_-]?key|access[_-]?token|refresh[_-]?token|token)\b\s*["']?\s*[:=]\s*)(?:(?:Bearer|Basic)\s+)?[^\s,;}\]"']+/gi, '$1[redacted]');
    safe = safe.replace(/([?&](?:token|access[_-]?token|api[_-]?key|authorization)=)[^&#\s"'<>]+/gi, '$1[redacted]');
    safe = safe.replace(/\b(?:https?|ftp):\/\/[^\s"'<>]+/gi, '[redacted URL]');
    return safe;
}

function safeMcpError(error, token) {
    const safe = new Error(sanitizeMcpErrorMessage(error && error.message, token));
    if (error && error.name) safe.name = sanitizeMcpErrorMessage(error.name, token);
    return safe;
}

let nextRequestId = 1;

/**
 * Performs one JSON-RPC POST and resolves the matching response message.
 *
 * @private
 */
async function rpcPost(url, { body, token, sessionId, fetchFn, timeoutMs, signal }) {
    const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
    let timer;
    let onAbort;
    let abortError;
    try {
        if (signal && signal.aborted) throw new DOMException('The operation was aborted.', 'AbortError');
        const cancelled = new Promise((_, reject) => {
            const abort = (message) => {
                abortError = new DOMException(message, 'AbortError');
                reject(abortError);
                if (controller) controller.abort();
            };
            onAbort = () => abort('The operation was aborted.');
            if (signal) signal.addEventListener('abort', onAbort, { once: true });
            timer = setTimeout(() => abort('MCP request timed out.'), timeoutMs);
        });
        const headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/event-stream',
        };
        if (token) headers['Authorization'] = `Bearer ${token}`;
        if (sessionId) headers['Mcp-Session-Id'] = sessionId;

        const response = await Promise.race([
            cancelled,
            Promise.resolve().then(() => {
                if (abortError) throw abortError;
                return fetchFn(url, {
                    method: 'POST',
                    headers,
                    body: JSON.stringify(body),
                    signal: controller ? controller.signal : undefined,
                });
            }),
        ]);
        if (abortError) throw abortError;

        if (!response.ok) {
            throw new Error(`MCP HTTP ${response.status}: ${response.statusText || 'request failed'}`);
        }
        const sessionIdFromServer = response.headers
            ? (response.headers.get('Mcp-Session-Id') || response.headers.get('mcp-session-id'))
            : null;

        const contentType = (response.headers && response.headers.get('content-type')) || '';
        let message = null;
        const text = await Promise.race([cancelled, response.text()]);
        if (abortError) throw abortError;
        if (contentType.includes('text/event-stream')) {
            for (const line of text.split(/\r?\n/)) {
                if (!line.startsWith('data:')) continue;
                const payload = line.slice(5).trim();
                if (!payload || payload === '[DONE]') continue;
                try {
                    const parsed = JSON.parse(payload);
                    if (body.id === undefined || parsed.id === body.id) {
                        message = parsed;
                        break;
                    }
                } catch (_parseErr) {
                    // Non-JSON data line — skip.
                }
            }
            if (!message) {
                // A notification-only stream is a valid reply to a notification.
                return { sessionId: sessionIdFromServer, message: null };
            }
        } else if (text) {
            try {
                message = JSON.parse(text);
            } catch (_parseErr) {
                // Empty or non-JSON body — valid for 202 notifications.
                return { sessionId: sessionIdFromServer, message: null };
            }
        }

        if (message && message.error) {
            const err = message.error;
            throw new Error(`MCP error ${err.code}: ${err.message || 'unknown error'}`);
        }
        return { sessionId: sessionIdFromServer, message };
    } catch (err) {
        throw safeMcpError(abortError || err, token);
    } finally {
        clearTimeout(timer);
        if (signal && onAbort) signal.removeEventListener('abort', onAbort);
    }
}

/**
 * Connects to an MCP server and returns a handle with typed helpers.
 * Each helper accepts a final options object whose signal overrides the default.
 *
 * @param {object} args
 * @param {string} args.url - MCP server endpoint (HTTPS, local HTTP, or same-origin relative path)
 * @param {string} [args.token] - Bearer token for the Authorization header
 * @param {function} [args.fetchFn] - Injectable fetch (defaults to global)
 * @param {function} [args.log] - Logging callback (message, type)
 * @param {number} [args.timeoutMs=30000] - Per-request deadline including response body
 * @param {AbortSignal} [args.signal] - Default signal for handshake and subsequent requests
 * @param {string} [args.protocolVersion=MCP_PROTOCOL_VERSION]
 * @returns {Promise<{serverInfo: {name: string, version: string}, protocolVersion: string,
 *   listTools: function({signal?: AbortSignal}=): Promise<Array>, callTool: function(string, object, {signal?: AbortSignal}=): Promise<object>,
 *   listPrompts: function({signal?: AbortSignal}=): Promise<Array>, getPrompt: function(string, object=, {signal?: AbortSignal}=): Promise<object>,
 *   listResources: function({signal?: AbortSignal}=): Promise<Array>, readResource: function(string, {signal?: AbortSignal}=): Promise<object>}>}
 * @throws {Error} On HTTP errors, JSON-RPC errors, cancellation, or handshake failures
 */
export async function connectMcpServer({
    url, token, fetchFn = (typeof fetch === 'function' ? fetch : null), log = () => {},
    timeoutMs = 30000, protocolVersion = MCP_PROTOCOL_VERSION, signal,
}) {
    if (!fetchFn) throw new Error('MCP client requires a fetch implementation');
    if (!url || typeof url !== 'string') throw new Error('MCP client requires a server URL');
    const endpoint = url.trim();
    if (!isAllowedMcpUrl(endpoint)) {
        throw new Error('MCP server URL must be an HTTPS URL, a localhost HTTP URL, or a relative path.');
    }

    let sessionId = null;

    const request = async (method, params, requestSignal = signal) => {
        const id = nextRequestId++;
        const { sessionId: returned, message } = await rpcPost(endpoint, {
            body: { jsonrpc: '2.0', id, method, params },
            token,
            sessionId,
            fetchFn,
            timeoutMs,
            signal: requestSignal,
        });
        if (returned) sessionId = returned;
        return message;
    };

    const notify = async (method, params) => {
        const { sessionId: returned } = await rpcPost(endpoint, {
            body: { jsonrpc: '2.0', method, params },
            token,
            sessionId,
            fetchFn,
            timeoutMs,
            signal,
        });
        if (returned) sessionId = returned;
    };

    const initMessage = await request('initialize', {
        protocolVersion,
        capabilities: {},
        clientInfo: { name: 'claric', version: '1.0.0' },
    });
    const result = (initMessage && initMessage.result) || {};
    const serverInfo = result.serverInfo || { name: 'unknown', version: '' };
    log(`Connected to MCP server "${serverInfo.name}" (protocol ${result.protocolVersion || protocolVersion}).`, 'info');

    await notify('notifications/initialized');

    return {
        serverInfo,
        protocolVersion: result.protocolVersion || protocolVersion,
        /** Lists the server's tools (MCP tools/list). */
        listTools: async ({ signal: requestSignal = signal } = {}) => {
            const message = await request('tools/list', {}, requestSignal);
            return (message && message.result && Array.isArray(message.result.tools)) ? message.result.tools : [];
        },
        /** Calls one tool (MCP tools/call); returns the raw result {content, isError}. */
        callTool: async (name, args, { signal: requestSignal = signal } = {}) => {
            const message = await request('tools/call', { name, arguments: args || {} }, requestSignal);
            return (message && message.result) || { content: [], isError: true };
        },
        /** Lists the server's prompt templates; servers without prompt
         *  support (JSON-RPC "Method not found") degrade to an empty list. */
        listPrompts: async ({ signal: requestSignal = signal } = {}) => {
            try {
                const message = await request('prompts/list', {}, requestSignal);
                return (message && message.result && Array.isArray(message.result.prompts)) ? message.result.prompts : [];
            } catch (err) {
                if (err.name !== 'AbortError' && /-32601|Method not found/i.test(err.message)) return [];
                throw err;
            }
        },
        /** Fetches one prompt template (MCP prompts/get); returns the raw result. */
        getPrompt: async (name, args, { signal: requestSignal = signal } = {}) => {
            const message = await request('prompts/get', { name, arguments: args || {} }, requestSignal);
            return (message && message.result) || { messages: [] };
        },
        /** Lists the server's resources; servers without resource support
         *  degrade to an empty list. */
        listResources: async ({ signal: requestSignal = signal } = {}) => {
            try {
                const message = await request('resources/list', {}, requestSignal);
                return (message && message.result && Array.isArray(message.result.resources)) ? message.result.resources : [];
            } catch (err) {
                if (err.name !== 'AbortError' && /-32601|Method not found/i.test(err.message)) return [];
                throw err;
            }
        },
        /** Reads one resource (MCP resources/read); returns the raw result. */
        readResource: async (uri, { signal: requestSignal = signal } = {}) => {
            const message = await request('resources/read', { uri }, requestSignal);
            return (message && message.result) || { contents: [] };
        },
    };
}
