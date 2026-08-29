/**
 * Production static file server for the Claric Word add-in.
 *
 * Serves the webpack build output (dist/) plus the generated manifest.xml
 * over HTTPS (default) or HTTP (PROTOCOL=http) for local testing, and
 * optionally proxies LLM API traffic on the same origin (/ollama, /vllm).
 *
 * LLM proxy routes are DISABLED BY DEFAULT in production: every
 * *_PROXY_PATH defaults to empty, and a provider's proxy is enabled only
 * by explicitly setting its *_PROXY_PATH env var (empty = disabled).
 *
 * Environment variables:
 *   PORT                  - listen port (default 3000)
 *   PROTOCOL              - 'https' (default) or 'http'
 *   SSL_CERT_FILE         - cert path, relative to project root or absolute
 *   SSL_KEY_FILE          - key path, relative to project root or absolute
 *   OLLAMA_PROXY_PATH     - proxy path for Ollama (default empty = disabled; set to e.g. /ollama to enable)
 *   OLLAMA_PROXY_TARGET   - upstream Ollama base URL (default http://localhost:11434)
 *   VLLM_PROXY_PATH       - proxy path for vLLM (default empty = disabled; set to e.g. /vllm to enable)
 *   VLLM_PROXY_TARGET     - upstream vLLM base URL (default http://localhost:8026)
 *   DEEPSEEK_PROXY_PATH   - proxy path for DeepSeek (default empty = disabled; set to e.g. /deepseek to enable)
 *   DEEPSEEK_PROXY_TARGET - upstream DeepSeek API origin (https://api.deepseek.com)
 *   GLM_PROXY_PATH        - proxy path for Zhipu GLM (default empty = disabled; set to e.g. /glm to enable)
 *   GLM_PROXY_TARGET      - upstream GLM API origin (https://open.bigmodel.cn)
 *   KIMI_PROXY_PATH       - proxy path for Moonshot Kimi (default empty = disabled; set to e.g. /kimi to enable)
 *   KIMI_PROXY_TARGET     - upstream Kimi API origin (https://api.moonshot.cn)
 *   MINIMAX_PROXY_PATH    - proxy path for MiniMax international (default empty = disabled; set to e.g. /minimax to enable)
 *   MINIMAX_PROXY_TARGET  - upstream MiniMax API origin (https://api.minimax.io)
 *   MINIMAX_CN_PROXY_PATH - proxy path for MiniMax China (default empty = disabled; set to e.g. /minimax-cn to enable)
 *   MINIMAX_CN_PROXY_TARGET - upstream MiniMax CN API origin (https://api.minimaxi.com)
 *   LLM_PROXY_TIMEOUT_MS  - upstream request timeout (default 300000 = 5 min)
 *
 * Why proxy LLM traffic at all: the add-in's UI defaults its backend URLs
 * to same-origin paths ('/ollama', '/vllm', '/deepseek', '/glm', '/kimi',
 * '/minimax', '/minimax-cn').
 * Serving those paths from the same HTTPS origin
 * avoids mixed-content blocking (https page fetching http://localhost) and
 * CORS configuration on the backend. This matters most when Word runs on
 * the same machine as the LLM (the common local-AI setup). Reach the host
 * from inside a Docker container via `host.docker.internal`.
 */

const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');
const { generateManifest } = require('./generate-manifest.cjs');
const { DEFAULT_LLM_PROXY_TIMEOUT_MS } = require('./llm-constants.cjs');

const rootDir = path.resolve(__dirname, '..');
const distDir = path.join(rootDir, 'dist');
const manifestPath = path.join(rootDir, 'manifest.xml');

// Requests are held open by Word until the page loads; give slow clients
// a bounded window instead of the 2-minute node default.
const REQUEST_TIMEOUT_MS = 60000;
const SHUTDOWN_TIMEOUT_MS = 10000;

// LLM proxy routes, built once in startServer from environment variables.
let PROXY_ROUTES = [];

function getEnv() {
  return {
    PORT: process.env.PORT || '3000',
    PROTOCOL: process.env.PROTOCOL || 'https',
    SSL_CERT_FILE: process.env.SSL_CERT_FILE || 'server.pem',
    SSL_KEY_FILE: process.env.SSL_KEY_FILE || 'server-key.pem',
    OLLAMA_PROXY_PATH: process.env.OLLAMA_PROXY_PATH || '',
    OLLAMA_PROXY_TARGET: process.env.OLLAMA_PROXY_TARGET || 'http://localhost:11434',
    VLLM_PROXY_PATH: process.env.VLLM_PROXY_PATH || '',
    VLLM_PROXY_TARGET: process.env.VLLM_PROXY_TARGET || 'http://localhost:8026',
    DEEPSEEK_PROXY_PATH: process.env.DEEPSEEK_PROXY_PATH || '',
    DEEPSEEK_PROXY_TARGET: process.env.DEEPSEEK_PROXY_TARGET || 'https://api.deepseek.com',
    GLM_PROXY_PATH: process.env.GLM_PROXY_PATH || '',
    GLM_PROXY_TARGET: process.env.GLM_PROXY_TARGET || 'https://open.bigmodel.cn',
    KIMI_PROXY_PATH: process.env.KIMI_PROXY_PATH || '',
    KIMI_PROXY_TARGET: process.env.KIMI_PROXY_TARGET || 'https://api.moonshot.cn',
    MINIMAX_PROXY_PATH: process.env.MINIMAX_PROXY_PATH || '',
    MINIMAX_PROXY_TARGET: process.env.MINIMAX_PROXY_TARGET || 'https://api.minimax.io',
    MINIMAX_CN_PROXY_PATH: process.env.MINIMAX_CN_PROXY_PATH || '',
    MINIMAX_CN_PROXY_TARGET: process.env.MINIMAX_CN_PROXY_TARGET || 'https://api.minimaxi.com',
    CUSTOM_PROXY_PATH: process.env.CUSTOM_PROXY_PATH || '',
    CUSTOM_PROXY_TARGET: process.env.CUSTOM_PROXY_TARGET || '',
    LLM_PROXY_TIMEOUT_MS: parseInt(process.env.LLM_PROXY_TIMEOUT_MS || String(DEFAULT_LLM_PROXY_TIMEOUT_MS), 10)
  };
}

// Headers that must not be forwarded to an upstream (RFC 7230 hop-by-hop).
const HOP_BY_HOP_HEADERS = new Set([
  'connection', 'keep-alive', 'proxy-authenticate', 'proxy-authorization',
  'proxy-connection', 'te', 'trailer', 'transfer-encoding', 'upgrade'
]);

// Browser-scope headers stripped before proxying: they carry the add-in
// origin's session state and are meaningless (at best) to the LLM upstream.
const STRIPPED_BROWSER_HEADERS = new Set(['cookie', 'referer', 'origin']);

function getContentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  switch (ext) {
    case '.html':
      return 'text/html; charset=utf-8';
    case '.js':
      return 'application/javascript; charset=utf-8';
    case '.css':
      return 'text/css; charset=utf-8';
    case '.json':
      return 'application/json; charset=utf-8';
    case '.png':
      return 'image/png';
    case '.jpg':
    case '.jpeg':
      return 'image/jpeg';
    case '.gif':
      return 'image/gif';
    case '.svg':
      return 'image/svg+xml';
    case '.ico':
      return 'image/x-icon';
    case '.map':
      return 'application/octet-stream';
    case '.xml':
      return 'application/xml; charset=utf-8';
    default:
      return 'application/octet-stream';
  }
}

/**
 * Resolves a URL path inside baseDir, rejecting traversal attempts.
 * The boundary check uses a path separator so sibling directories that
 * merely share a string prefix (e.g. /app/dist vs /app/dist-backup)
 * are rejected too.
 */
function safeJoin(baseDir, targetPath) {
  const baseWithSep = baseDir.endsWith(path.sep) ? baseDir : baseDir + path.sep;
  const resolvedPath = path.resolve(baseDir, `.${targetPath}`);
  if (resolvedPath !== baseDir && !resolvedPath.startsWith(baseWithSep)) {
    return null;
  }
  return resolvedPath;
}

/**
 * Sends a plain-text error response.
 */
function sendError(res, statusCode, reason) {
  res.writeHead(statusCode, {
    'Content-Type': 'text/plain; charset=utf-8',
    'Cache-Control': 'no-store'
  });
  res.end(reason);
}

function serveFile(res, filePath) {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      // EISDIR (directory requested) and ENOENT (missing) both map to 404;
      // anything else is a server-side problem worth surfacing.
      const notFound = err.code === 'ENOENT' || err.code === 'EISDIR';
      if (!notFound) {
        console.error(`Error reading ${filePath}:`, err.message);
      }
      sendError(res, notFound ? 404 : 500, notFound ? 'Not found' : 'Internal server error');
      return;
    }
    res.writeHead(200, {
      'Content-Type': getContentType(filePath),
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
      'Referrer-Policy': 'no-referrer'
    });
    res.end(data);
  });
}

/**
 * Builds LLM proxy routes from environment variables.
 *
 * A route is created only when both a path and a reachable http(s) target
 * are configured; an empty path disables that backend entirely.
 *
 * @param {object} env - Result of getEnv()
 * @returns {Array<{proxyPath: string, targetUrl: URL, timeoutMs: number, agent: object|null}>}
 */
function buildProxyRoutes(env) {
  const routes = [];
  const backends = [
    ['OLLAMA_PROXY_PATH', 'OLLAMA_PROXY_TARGET'],
    ['VLLM_PROXY_PATH', 'VLLM_PROXY_TARGET'],
    ['DEEPSEEK_PROXY_PATH', 'DEEPSEEK_PROXY_TARGET'],
    ['GLM_PROXY_PATH', 'GLM_PROXY_TARGET'],
    ['KIMI_PROXY_PATH', 'KIMI_PROXY_TARGET'],
    ['MINIMAX_PROXY_PATH', 'MINIMAX_PROXY_TARGET'],
    ['MINIMAX_CN_PROXY_PATH', 'MINIMAX_CN_PROXY_TARGET'],
    ['CUSTOM_PROXY_PATH', 'CUSTOM_PROXY_TARGET'],
  ];

  for (const [pathKey, targetKey] of backends) {
    const proxyPath = env[pathKey];
    const target = env[targetKey];
    if (!proxyPath || !target) continue;

    let targetUrl;
    try {
      targetUrl = new URL(target);
    } catch {
      console.error(`Ignoring ${targetKey} (invalid URL): ${target}`);
      continue;
    }
    if (targetUrl.protocol !== 'http:' && targetUrl.protocol !== 'https:') {
      console.error(`Ignoring ${targetKey} (must be http or https): ${target}`);
      continue;
    }

    routes.push({
      proxyPath: proxyPath.replace(/\/+$/, ''),
      targetUrl,
      timeoutMs: env.LLM_PROXY_TIMEOUT_MS,
      agent: null,
    });
  }
  return routes;
}

/**
 * Proxies an API request to the configured upstream LLM backend.
 *
 * The path suffix after the proxy prefix (including the query string) is
 * forwarded verbatim. Response status/headers stream back as-is; failures
 * before the response starts yield 502, timeouts yield 504.
 *
 * @param {object} route - Entry from buildProxyRoutes()
 * @param {http.IncomingMessage} req
 * @param {http.ServerResponse} res
 * @param {string} urlPath - Decoded request path (no query string)
 */
function handleProxyRequest(route, req, res, urlPath) {
  // Same-origin from the add-in's perspective, but the WebView may still
  // issue a preflight; answer it like the dev-server proxy does.
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, PUT, PATCH, DELETE, HEAD, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization, Accept',
      'Access-Control-Max-Age': '86400',
    });
    res.end();
    return;
  }

  const query = req.url.split('?')[1];
  const upstreamPath = urlPath.slice(route.proxyPath.length) + (query ? `?${query}` : '');
  const upstream = new URL(upstreamPath, route.targetUrl);
  // `new URL(suffix, base)` rebases onto a protocol-relative suffix ("//host"
  // or "/\\host" after the prefix is stripped), letting a request path pick
  // an arbitrary authority — an SSRF hole. The suffix may only select a path
  // on the configured target's own origin; anything else is rejected before
  // a socket is opened.
  if (upstream.origin !== route.targetUrl.origin) {
    console.error(`[${route.proxyPath} Proxy] rejected cross-origin path: ${urlPath}`);
    sendError(res, 400, 'Invalid proxy path');
    return;
  }

  const headers = {};
  for (const [name, value] of Object.entries(req.headers)) {
    if (HOP_BY_HOP_HEADERS.has(name)) continue;
    // Browser-scope headers that carry same-origin session state must not
    // leak to the LLM upstream. (Authorization — the API key — IS forwarded.)
    if (STRIPPED_BROWSER_HEADERS.has(name)) continue;
    headers[name] = value;
  }
  headers.host = route.targetUrl.host;

  if (!route.agent) {
    const AgentCtor = route.targetUrl.protocol === 'https:' ? https.Agent : http.Agent;
    route.agent = new AgentCtor({ keepAlive: true, maxSockets: 50, maxFreeSockets: 10 });
  }

  const client = route.targetUrl.protocol === 'https:' ? https : http;
  let timedOut = false;

  const proxyReq = client.request(
    {
      method: req.method,
      hostname: upstream.hostname,
      port: upstream.port || (upstream.protocol === 'https:' ? 443 : 80),
      path: upstream.pathname + upstream.search,
      headers,
      agent: route.agent,
    },
    (proxyRes) => {
      res.writeHead(proxyRes.statusCode, proxyRes.headers);
      proxyRes.pipe(res);
    }
  );

  proxyReq.setTimeout(route.timeoutMs, () => {
    timedOut = true;
    proxyReq.destroy();
  });

  proxyReq.on('error', (err) => {
    const reason = err.message || err.code || 'unknown upstream error';
    console.error(`[${route.proxyPath} Proxy Error]`, reason);
    if (!res.headersSent) {
      sendError(res, timedOut ? 504 : 502, timedOut ? 'LLM upstream timed out' : `LLM proxy error: ${reason}`);
    } else {
      res.end();
    }
  });

  req.pipe(proxyReq);
}

/**
 * Handles one request. May throw — requestHandler wraps this so an
 * unexpected synchronous throw answers 500 on that connection instead of
 * escaping as an uncaught exception and killing the process.
 */
function handleRequest(req, res) {
  let urlPath;
  try {
    urlPath = decodeURIComponent(req.url.split('?')[0]);
  } catch {
    // Malformed percent-encoding (e.g. GET /%) -- reject, never crash.
    sendError(res, 400, 'Bad request');
    return;
  }

  res.on('finish', () => {
    console.log(`${new Date().toISOString()} ${req.method} ${req.url} -> ${res.statusCode}`);
  });

  // A decoded control character (NUL, newline, ...) never occurs in a
  // legitimate asset or API path, and NUL is worse than invalid: fs.readFile
  // throws synchronously on it (ERR_INVALID_ARG_VALUE), so one crafted
  // request like GET /%00 used to take the whole process down. Reject
  // before anything touches the path.
  // eslint-disable-next-line no-control-regex -- matching control characters is the purpose of this validation
  if (/[\u0000-\u001f\u007f]/.test(urlPath)) {
    sendError(res, 400, 'Bad request');
    return;
  }

  if (urlPath === '/healthz') {
    res.writeHead(200, {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    });
    res.end(JSON.stringify({ status: 'ok' }));
    return;
  }

  // LLM proxy routes are matched before the GET/HEAD static restriction:
  // chat completions are POST, and the add-in talks to '/ollama'/'/vllm'
  // on this same origin.
  const proxyRoute = PROXY_ROUTES.find(
    (route) => urlPath === route.proxyPath || urlPath.startsWith(route.proxyPath + '/')
  );
  if (proxyRoute) {
    handleProxyRequest(proxyRoute, req, res, urlPath);
    return;
  }

  if (req.method !== 'GET' && req.method !== 'HEAD') {
    res.setHeader('Allow', 'GET, HEAD');
    sendError(res, 405, 'Method not allowed');
    return;
  }

  if (urlPath === '/' || urlPath === '') {
    serveFile(res, path.join(distDir, 'taskpane.html'));
    return;
  }

  if (urlPath === '/manifest.xml') {
    serveFile(res, manifestPath);
    return;
  }

  const filePath = safeJoin(distDir, urlPath);
  if (!filePath) {
    sendError(res, 400, 'Bad request');
    return;
  }

  serveFile(res, filePath);
}

/**
 * Crash-proof entry point wired into the HTTP server: an unexpected
 * synchronous throw inside request handling answers 500 on that connection
 * instead of killing the server — and every in-flight LLM stream with it.
 */
function requestHandler(req, res) {
  try {
    handleRequest(req, res);
  } catch (err) {
    console.error('Request handling error:', (err && err.message) || err);
    if (!res.headersSent) {
      sendError(res, 500, 'Internal server error');
    } else {
      res.end();
    }
  }
}

/**
 * Wires graceful shutdown and top-level error handling on the server.
 *
 * @param {http.Server|https.Server} server
 * @returns {http.Server|https.Server} the same server, for chaining
 */
function configureServer(server) {
  server.headersTimeout = REQUEST_TIMEOUT_MS + 5000;
  server.requestTimeout = REQUEST_TIMEOUT_MS;

  // Malformed HTTP from the network lands here -- respond 400 instead of
  // letting node emit an uncaught 'clientError'.
  server.on('clientError', (err, socket) => {
    if (socket.writable) {
      socket.end('HTTP/1.1 400 Bad Request\r\n\r\n');
    }
  });

  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.error(`Port ${getEnv().PORT} is already in use. Exiting.`);
    } else {
      console.error('Server error:', err.message);
    }
    process.exit(1);
  });

  // Last-resort process guards. This server is stateless and each request
  // is independent (LLM streams can run for minutes), so an isolated
  // asynchronous throw costs less as a logged-and-dropped connection than
  // as a process exit that kills every in-flight stream. requestHandler
  // already catches throws at the request boundary; these catch whatever
  // escapes it.
  process.on('uncaughtException', (err) => {
    console.error('Uncaught exception (server kept alive):', (err && err.stack) || err);
  });
  process.on('unhandledRejection', (reason) => {
    console.error('Unhandled rejection (server kept alive):', (reason && reason.stack) || reason);
  });

  const shutdown = (signal) => {
    console.log(`\n${signal} received -- shutting down (waiting up to ${SHUTDOWN_TIMEOUT_MS / 1000}s)...`);
    server.close(() => {
      console.log('All connections closed.');
      process.exit(0);
    });
    // In-flight requests get a grace window; then exit anyway.
    setTimeout(() => process.exit(0), SHUTDOWN_TIMEOUT_MS).unref();
  };
  process.on('SIGINT', () => shutdown('SIGINT'));
  process.on('SIGTERM', () => shutdown('SIGTERM'));

  return server;
}

function startServer() {
  if (!fs.existsSync(distDir)) {
    console.error('Missing dist/ directory. Did the build complete?');
    process.exit(1);
  }

  try {
    generateManifest({ rootDir });
  } catch (err) {
    console.error(`Manifest generation failed: ${err.message}`);
    process.exit(1);
  }

  const env = getEnv();
  PROXY_ROUTES = buildProxyRoutes(env);
  if (PROXY_ROUTES.length > 0) {
    console.log(
      'LLM proxy: ' +
      PROXY_ROUTES.map((r) => `${r.proxyPath} -> ${r.targetUrl.href} (${r.timeoutMs}ms timeout)`).join(', ')
    );
  }
  const port = Number(env.PORT);

  if (env.PROTOCOL === 'http') {
    console.warn('WARNING: serving over plain HTTP. Word requires HTTPS for add-in hosting; use HTTP only for local testing.');
    configureServer(http.createServer(requestHandler)).listen(port, () => {
      console.log(`HTTP server running on port ${port}`);
    });
    return;
  }

  const certPath = path.isAbsolute(env.SSL_CERT_FILE)
    ? env.SSL_CERT_FILE
    : path.join(rootDir, env.SSL_CERT_FILE);
  const keyPath = path.isAbsolute(env.SSL_KEY_FILE)
    ? env.SSL_KEY_FILE
    : path.join(rootDir, env.SSL_KEY_FILE);

  if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
    console.error('Missing SSL cert/key files. Provide server.pem and server-key.pem.');
    process.exit(1);
  }

  configureServer(
    https.createServer(
      {
        cert: fs.readFileSync(certPath),
        key: fs.readFileSync(keyPath)
      },
      requestHandler
    )
  ).listen(port, () => {
    console.log(`HTTPS server running on port ${port}`);
  });
}

// Run directly (`node scripts/docker-server.cjs`): boot the server.
// Required as a module (tests): export the pieces worth unit-testing and
// leave the process alone.
if (require.main === module) {
  startServer();
}

module.exports = { buildProxyRoutes, handleProxyRequest, requestHandler, configureServer };
