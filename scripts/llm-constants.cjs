/**
 * Shared constants for the dev-server (webpack.config.cjs) and production
 * (docker-server.cjs) LLM proxies. LLM calls can take minutes, so both
 * timeouts are 5 minutes.
 */

const DEFAULT_LLM_PROXY_TIMEOUT_MS = 300000;

module.exports = { DEFAULT_LLM_PROXY_TIMEOUT_MS };
