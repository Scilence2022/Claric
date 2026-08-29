/**
 * MCP Tools Bridge
 *
 * Converts MCP servers' tools into the Claric tool-loop's ReAct specs and
 * dispatches loop tool calls back to the right MCP server. This is the L1
 * seam: lib/mcp-client.js speaks MCP, lib/tool-loop.js speaks one-JSON-call
 * per turn — this module is the contract between them.
 *
 * Loop tool names must match ^[a-z][a-z0-9_]*$ and be unique across all
 * servers, so MCP tool names are sanitized and namespaced on collision.
 *
 * Security contract: MCP tool results are OBSERVATIONS (data), never
 * direct document writes — anything that should mutate the Word document
 * still has to go through the proposal-card staging pipeline.
 *
 * Pure module — no DOM, no network (the clients carry it).
 *
 * @module mcp-tools
 */

import { defineTool } from './tool-registry.js';

/** Max characters of one MCP tool result accepted into an observation. */
export const MAX_MCP_RESULT_CHARS = 64 * 1024;

/**
 * Sanitizes an MCP tool name into the loop's snake_case namespace.
 *
 * @param {string} raw
 * @returns {string}
 * @private
 */
function sanitizeToolName(raw) {
    const cleaned = String(raw || '').toLowerCase().replace(/[^a-z0-9_]/g, '_');
    return /^[a-z]/.test(cleaned) ? cleaned : `mcp_${cleaned}`;
}

/**
 * Derives a one-level example args object from a JSON Schema (examples
 * teach the ReAct loop's small-model protocol better than schemas).
 *
 * @param {object} [schema] - MCP inputSchema (JSON Schema)
 * @returns {object}
 * @private
 */
function exampleFromSchema(schema) {
    const example = {};
    const props = schema && schema.properties && typeof schema.properties === 'object'
        ? schema.properties : {};
    for (const [key, prop] of Object.entries(props)) {
        switch (prop && prop.type) {
            case 'string': example[key] = prop.enum ? prop.enum[0] : '...'; break;
            case 'number':
            case 'integer': example[key] = 0; break;
            case 'boolean': example[key] = true; break;
            case 'array': example[key] = []; break;
            case 'object': example[key] = {}; break;
            default: example[key] = '...';
        }
    }
    return example;
}

/**
 * Builds the loop-side tool specs and the call→server mapping for a set of
 * connected MCP clients.
 *
 * @param {Array<{name: string, client: object}>} clients - Connected clients
 *   paired with their mcpTools arrays (from client.listTools()).
 * @param {Array<{name: string, client: object, mcpTools: Array}>} _clients
 * @returns {{loopTools: Array, mapping: Map<string, {client: object, originalName: string, serverName: string}>}}
 */
export function buildLoopTools(clients) {
    const loopTools = [];
    const mapping = new Map();
    const seen = new Set();

    for (const entry of Array.isArray(clients) ? clients : []) {
        const serverName = String(entry.name || 'server');
        for (const tool of Array.isArray(entry.mcpTools) ? entry.mcpTools : []) {
            if (!tool || typeof tool.name !== 'string') continue;
            let loopName = sanitizeToolName(tool.name);
            if (seen.has(loopName)) loopName = `${loopName}_${serverName.replace(/[^a-z0-9_]/gi, '_').toLowerCase()}`;
            if (seen.has(loopName)) continue; // still colliding — skip rather than misroute
            seen.add(loopName);

            const description = `[MCP:${serverName}] ${tool.description || tool.name}`;
            loopTools.push(defineTool({
                name: loopName,
                description,
                argsExample: exampleFromSchema(tool.inputSchema),
            }));
            mapping.set(loopName, { client: entry.client, originalName: tool.name, serverName });
        }
    }
    return { loopTools, mapping };
}

/**
 * Converts an MCP tools/call result into a tool-loop observation.
 * Text content blocks join into the result string; image content blocks
 * ride the loop's attachments channel (data URLs) so vision-capable
 * backends can see them.
 *
 * @param {object} mcpResult - Raw MCP result {content: Array, isError?: boolean}
 * @param {object} [options]
 * @param {number} [options.maxChars=MAX_MCP_RESULT_CHARS]
 * @returns {{ok: boolean, result?: string, error?: string, attachments?: Array<{dataUrl: string}>}}
 * @private
 */
function observationFromResult(mcpResult, { maxChars = MAX_MCP_RESULT_CHARS } = {}) {
    const blocks = Array.isArray(mcpResult && mcpResult.content) ? mcpResult.content : [];
    const texts = [];
    const attachments = [];
    for (const block of blocks) {
        if (block && block.type === 'text' && typeof block.text === 'string') {
            texts.push(block.text);
        } else if (block && block.type === 'image' && typeof block.data === 'string') {
            const mime = block.mimeType || 'image/png';
            attachments.push({ dataUrl: `data:${mime};base64,${block.data}` });
        }
    }
    let text = texts.join('\n');
    if (text.length > maxChars) {
        text = `${text.slice(0, maxChars)}\n… [truncated ${text.length - maxChars} chars]`;
    }
    if (mcpResult && mcpResult.isError) {
        return { ok: false, error: text || 'The MCP tool reported an error.' };
    }
    const observation = { ok: true, result: text };
    if (attachments.length > 0) observation.attachments = attachments;
    return observation;
}

/**
 * Creates the tool-loop executor that routes loop tool calls to MCP
 * servers. Failures become error observations (the loop feeds them back to
 * the model) — only transport-level surprises would throw.
 *
 * @param {Map<string, {client: object, originalName: string, serverName: string}>} mapping
 * @param {object} [options]
 * @param {number} [options.maxChars=MAX_MCP_RESULT_CHARS]
 * @returns {function(string, object): Promise<object>} execute(toolName, args) → observation
 */
export function createMcpToolExecutor(mapping, { maxChars = MAX_MCP_RESULT_CHARS } = {}) {
    return async function execute(toolName, args) {
        const entry = mapping.get(toolName);
        if (!entry) {
            return { ok: false, error: `Unknown MCP tool "${toolName}".` };
        }
        try {
            const result = await entry.client.callTool(entry.originalName, args);
            const observation = observationFromResult(result, { maxChars });
            if (observation.ok === false) {
                return { ok: false, error: observation.error };
            }
            return observation;
        } catch (err) {
            return { ok: false, error: `MCP call failed: ${err.message}` };
        }
    };
}
