/**
 * MCP Client (Streamable HTTP)
 *
 * Minimal Model Context Protocol client for the browser/WebView2 taskpane:
 * JSON-RPC over the Streamable HTTP transport (2025-06-18 protocol).
 * Browser sandboxes have no stdio, so HTTP is the only transport this
 * add-in can speak; CORS/mixed-content are handled by the same-origin
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

let nextRequestId = 1;

/**
 * Performs one JSON-RPC POST and resolves the matching response message.
 *
 * @private
 */
async function rpcPost(url, { body, token, sessionId, fetchFn, timeoutMs }) {
    const headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/event-stream',
    };
    if (token) headers['Authorization'] = `Bearer ${token}`;
    if (sessionId) headers['Mcp-Session-Id'] = sessionId;

    const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
    const timer = setTimeout(() => controller && controller.abort(), timeoutMs);
    let response;
    try {
        response = await fetchFn(url, {
            method: 'POST',
            headers,
            body: JSON.stringify(body),
            signal: controller ? controller.signal : undefined,
        });
    } finally {
        clearTimeout(timer);
    }

    if (!response.ok) {
        throw new Error(`MCP HTTP ${response.status}: ${response.statusText || 'request failed'}`);
    }
    const sessionIdFromServer = response.headers
        ? (response.headers.get('Mcp-Session-Id') || response.headers.get('mcp-session-id'))
        : null;

    const contentType = (response.headers && response.headers.get('content-type')) || '';
    let message = null;
    const text = await response.text();
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
}

/**
 * Connects to an MCP server and returns a handle with typed helpers.
 *
 * @param {object} args
 * @param {string} args.url - MCP server endpoint (HTTP; same-origin proxy path or absolute URL)
 * @param {string} [args.token] - Bearer token for the Authorization header
 * @param {function} [args.fetchFn] - Injectable fetch (defaults to global)
 * @param {function} [args.log] - Logging callback (message, type)
 * @param {number} [args.timeoutMs=30000] - Per-request timeout
 * @param {string} [args.protocolVersion=MCP_PROTOCOL_VERSION]
 * @returns {Promise<{serverInfo: {name: string, version: string}, protocolVersion: string,
 *   listTools: function(): Promise<Array>, callTool: function(string, object): Promise<object>,
 *   listPrompts: function(): Promise<Array>, getPrompt: function(string, object=): Promise<object>,
 *   listResources: function(): Promise<Array>, readResource: function(string): Promise<object>}>}
 * @throws {Error} On HTTP errors, JSON-RPC errors, or handshake failures
 */
export async function connectMcpServer({
    url, token, fetchFn = (typeof fetch === 'function' ? fetch : null), log = () => {},
    timeoutMs = 30000, protocolVersion = MCP_PROTOCOL_VERSION,
}) {
    if (!fetchFn) throw new Error('MCP client requires a fetch implementation');
    if (!url || typeof url !== 'string') throw new Error('MCP client requires a server URL');

    let sessionId = null;

    const request = async (method, params) => {
        const id = nextRequestId++;
        const { sessionId: returned, message } = await rpcPost(url, {
            body: { jsonrpc: '2.0', id, method, params },
            token,
            sessionId,
            fetchFn,
            timeoutMs,
        });
        if (returned) sessionId = returned;
        return message;
    };

    const notify = async (method, params) => {
        const { sessionId: returned } = await rpcPost(url, {
            body: { jsonrpc: '2.0', method, params },
            token,
            sessionId,
            fetchFn,
            timeoutMs,
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
        listTools: async () => {
            const message = await request('tools/list', {});
            return (message && message.result && Array.isArray(message.result.tools)) ? message.result.tools : [];
        },
        /** Calls one tool (MCP tools/call); returns the raw result {content, isError}. */
        callTool: async (name, args) => {
            const message = await request('tools/call', { name, arguments: args || {} });
            return (message && message.result) || { content: [], isError: true };
        },
        /** Lists the server's prompt templates; servers without prompt
         *  support (JSON-RPC "Method not found") degrade to an empty list. */
        listPrompts: async () => {
            try {
                const message = await request('prompts/list', {});
                return (message && message.result && Array.isArray(message.result.prompts)) ? message.result.prompts : [];
            } catch (err) {
                if (/-32601|Method not found/i.test(err.message)) return [];
                throw err;
            }
        },
        /** Fetches one prompt template (MCP prompts/get); returns the raw result. */
        getPrompt: async (name, args) => {
            const message = await request('prompts/get', { name, arguments: args || {} });
            return (message && message.result) || { messages: [] };
        },
        /** Lists the server's resources; servers without resource support
         *  degrade to an empty list. */
        listResources: async () => {
            try {
                const message = await request('resources/list', {});
                return (message && message.result && Array.isArray(message.result.resources)) ? message.result.resources : [];
            } catch (err) {
                if (/-32601|Method not found/i.test(err.message)) return [];
                throw err;
            }
        },
        /** Reads one resource (MCP resources/read); returns the raw result. */
        readResource: async (uri) => {
            const message = await request('resources/read', { uri });
            return (message && message.result) || { contents: [] };
        },
    };
}
