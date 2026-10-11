/**
 * The MCP surface the /mcp turn needs, as one lazily loaded chunk: MCP is
 * opt-in and rare, so its client and bridge stay out of the taskpane entry.
 * Re-exports keep live bindings to the source modules.
 *
 * @module mcp-runtime
 */
export { connectMcpServer } from './mcp-client.js';
export { buildLoopTools, createMcpToolExecutor, createResourceClient, RESOURCE_TOOL_SPECS, MCP_APPROVAL } from './mcp-tools.js';
