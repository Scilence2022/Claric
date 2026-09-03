const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');
const fs = require('fs');

// Load environment variables from .env file
require('dotenv').config();

const DEFAULT_DEV_SERVER_ALLOWED_HOSTS = Object.freeze([
  'localhost',
  '127.0.0.1',
  '[::1]'
]);
const CORS_ALLOWED_METHODS = 'GET, POST, PUT, DELETE, OPTIONS';
const CORS_ALLOWED_HEADERS = 'Content-Type, Authorization, Accept, X-Requested-With, x-api-key, anthropic-version, anthropic-dangerous-direct-browser-access';

/**
 * Parses the dev-server host allowlist. The special `all` value is never
 * accepted here: callers must opt into each additional host explicitly.
 *
 * @param {string|undefined} value
 * @returns {string[]}
 */
function parseAllowedHosts(value) {
  const hosts = String(value || '')
    .split(',')
    .map((host) => host.trim())
    .filter((host) => host && host.toLowerCase() !== 'all');
  return hosts.length > 0 ? hosts : [...DEFAULT_DEV_SERVER_ALLOWED_HOSTS];
}

/**
 * Normalizes one configured CORS origin. `*` remains available as an explicit
 * opt-in for test harnesses; malformed or non-HTTP(S) origins are ignored.
 *
 * @param {string|undefined} value
 * @returns {string|null}
 */
function normalizeCorsOrigin(value) {
  const candidate = String(value || '').trim();
  if (candidate === '*') return '*';
  if (!candidate) return null;
  if (!/^https?:\/\//i.test(candidate)) return null;
  try {
    const parsed = new URL(candidate);
    if (!parsed.hostname || parsed.username || parsed.password) return null;
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return null;
    return parsed.origin;
  } catch (_err) {
    return null;
  }
}

/**
 * Returns origins allowed by default for the HTTPS local dev server.
 *
 * @param {number} port
 * @returns {string[]}
 */
function defaultCorsOrigins(port) {
  return [
    `https://localhost:${port}`,
    `https://127.0.0.1:${port}`,
    `https://[::1]:${port}`,
  ].map(normalizeCorsOrigin).filter(Boolean);
}

/**
 * Parses DEV_SERVER_CORS_ORIGIN as a comma-separated origin list.
 *
 * @param {string|undefined} value
 * @param {number} port
 * @returns {string[]}
 */
function parseCorsOrigins(value, port) {
  const configured = String(value || '')
    .split(',')
    .map(normalizeCorsOrigin)
    .filter(Boolean);
  return configured.length > 0 ? configured : defaultCorsOrigins(port);
}

/**
 * Selects a CORS origin without reflecting an untrusted Origin header.
 *
 * @param {string|undefined} requestOrigin
 * @param {string[]} allowedOrigins
 * @returns {string|null}
 */
function resolveCorsOrigin(requestOrigin, allowedOrigins) {
  if (allowedOrigins.includes('*')) return '*';
  const normalizedRequestOrigin = normalizeCorsOrigin(requestOrigin);
  if (normalizedRequestOrigin && allowedOrigins.includes(normalizedRequestOrigin)) {
    return normalizedRequestOrigin;
  }
  return requestOrigin ? null : (allowedOrigins[0] || null);
}

/**
 * Builds the CORS headers used by static dev-server responses and proxy hooks.
 *
 * @param {object} req
 * @param {string[]} allowedOrigins
 * @returns {object}
 */
function corsHeaders(req, allowedOrigins) {
  const origin = resolveCorsOrigin(req && req.headers && req.headers.origin, allowedOrigins);
  const headers = {
    'Access-Control-Allow-Methods': CORS_ALLOWED_METHODS,
    'Access-Control-Allow-Headers': CORS_ALLOWED_HEADERS,
  };
  if (origin) {
    headers['Access-Control-Allow-Origin'] = origin;
    if (origin !== '*') headers.Vary = 'Origin';
  }
  return headers;
}

const LOOPBACK_PROXY_HOSTS = new Set([
  'localhost',
  '127.0.0.1',
  '[::1]',
  'host.docker.internal'
]);

/**
 * Accepts HTTPS proxy targets plus HTTP targets that resolve to loopback or
 * Docker's explicit host bridge alias.
 *
 * @param {string} value
 * @returns {URL|null}
 */
function parseProxyTarget(value) {
  const candidate = typeof value === 'string' ? value.trim() : '';
  if (!/^https?:\/\//i.test(candidate) || candidate.includes('\\')) return null;
  let parsed;
  try {
    parsed = new URL(candidate);
  } catch (_err) {
    return null;
  }
  if (!parsed.hostname || parsed.username || parsed.password) return null;
  if (parsed.protocol === 'https:') return parsed;
  if (parsed.protocol !== 'http:') return null;
  const hostname = parsed.hostname.toLowerCase().replace(/\.$/, '');
  return LOOPBACK_PROXY_HOSTS.has(hostname) ? parsed : null;
}

// Environment configuration with defaults
const ENV = {
  // Dev server (binds localhost by default; set DEV_SERVER_HOST=0.0.0.0
  // explicitly if the add-in must be reached from another machine)
  DEV_SERVER_HOST: process.env.DEV_SERVER_HOST || '127.0.0.1',
  DEV_SERVER_PORT: parseInt(process.env.DEV_SERVER_PORT || '3000', 10),
  DEV_SERVER_ALLOWED_HOSTS: parseAllowedHosts(process.env.DEV_SERVER_ALLOWED_HOSTS),
  DEV_SERVER_CORS_ORIGINS: parseCorsOrigins(
    process.env.DEV_SERVER_CORS_ORIGIN,
    parseInt(process.env.DEV_SERVER_PORT || '3000', 10)
  ),
  // Ollama proxy
  OLLAMA_PROXY_PATH: process.env.OLLAMA_PROXY_PATH || '/ollama',
  OLLAMA_PROXY_TARGET: process.env.OLLAMA_PROXY_TARGET || 'http://localhost:11434',
  // UI defaults (injected into bundle)
  DEFAULT_OLLAMA_URL: process.env.DEFAULT_OLLAMA_URL || '/ollama',
  DEFAULT_MODEL: process.env.DEFAULT_MODEL || 'gpt-oss:20b',
  // vLLM proxy
  VLLM_PROXY_PATH: process.env.VLLM_PROXY_PATH || '/vllm',
  VLLM_PROXY_TARGET: process.env.VLLM_PROXY_TARGET || 'http://localhost:8026',
  // vLLM UI defaults (injected into bundle)
  DEFAULT_VLLM_URL: process.env.DEFAULT_VLLM_URL || '/vllm',
  DEFAULT_VLLM_MODEL: process.env.VLLM_MODEL || 'qwen3.5-35b-a3b',
  // Cloud provider proxies (same-origin paths for the add-in)
  OPENAI_PROXY_PATH: process.env.OPENAI_PROXY_PATH || '/openai',
  OPENAI_PROXY_TARGET: process.env.OPENAI_PROXY_TARGET || 'https://api.openai.com',
  CLAUDE_PROXY_PATH: process.env.CLAUDE_PROXY_PATH || '/claude',
  CLAUDE_PROXY_TARGET: process.env.CLAUDE_PROXY_TARGET || 'https://api.anthropic.com',
  DEEPSEEK_PROXY_PATH: process.env.DEEPSEEK_PROXY_PATH || '/deepseek',
  DEEPSEEK_PROXY_TARGET: process.env.DEEPSEEK_PROXY_TARGET || 'https://api.deepseek.com',
  GLM_PROXY_PATH: process.env.GLM_PROXY_PATH || '/glm',
  GLM_PROXY_TARGET: process.env.GLM_PROXY_TARGET || 'https://open.bigmodel.cn',
  KIMI_PROXY_PATH: process.env.KIMI_PROXY_PATH || '/kimi',
  KIMI_PROXY_TARGET: process.env.KIMI_PROXY_TARGET || 'https://api.moonshot.cn',
  MINIMAX_PROXY_PATH: process.env.MINIMAX_PROXY_PATH || '/minimax',
  MINIMAX_PROXY_TARGET: process.env.MINIMAX_PROXY_TARGET || 'https://api.minimax.io',
  MINIMAX_CN_PROXY_PATH: process.env.MINIMAX_CN_PROXY_PATH || '/minimax-cn',
  MINIMAX_CN_PROXY_TARGET: process.env.MINIMAX_CN_PROXY_TARGET || 'https://api.minimaxi.com',
  ZHONGKEYU_PROXY_PATH: process.env.ZHONGKEYU_PROXY_PATH || '/zhongkeyu',
  ZHONGKEYU_PROXY_TARGET: process.env.ZHONGKEYU_PROXY_TARGET || 'https://zhongkeyu.com',
  // Custom (any OpenAI-compatible endpoint). Path defaults to empty so the
  // dev server does not register a route unless the user opts in.
  CUSTOM_PROXY_PATH: process.env.CUSTOM_PROXY_PATH || '',
  CUSTOM_PROXY_TARGET: process.env.CUSTOM_PROXY_TARGET || '',
};


/**
 * Builds http-proxy-middleware entries for every enabled LLM provider.
 *
 * Each proxy strips its path prefix, rewrites Origin/Referer off the
 * upstream request (some backends enforce CORS on the server side),
 * answers OPTIONS preflights directly, and adds CORS headers only for
 * configured local origins. LLM calls can take minutes, so both timeouts are
 * 5 minutes and connections use a keep-alive agent.
 *
 * A provider is disabled by setting its *_PROXY_PATH to an empty string.
 *
 * @param {object} ENV - Parsed environment configuration
 * @returns {Array<object>} Proxy config array for webpack-dev-server (v5 array form)
 */
function buildLlmProxies(ENV) {
  const LLM_PROXY_TIMEOUT_MS = require('./scripts/llm-constants.cjs').DEFAULT_LLM_PROXY_TIMEOUT_MS;
  const allowedCorsOrigins = Array.isArray(ENV.DEV_SERVER_CORS_ORIGINS)
    ? ENV.DEV_SERVER_CORS_ORIGINS
    : defaultCorsOrigins(ENV.DEV_SERVER_PORT || 3000);

  const providers = [
    ['OLLAMA_PROXY_PATH', 'OLLAMA_PROXY_TARGET', 'Ollama'],
    ['VLLM_PROXY_PATH', 'VLLM_PROXY_TARGET', 'vLLM'],
    ['OPENAI_PROXY_PATH', 'OPENAI_PROXY_TARGET', 'OpenAI'],
    ['CLAUDE_PROXY_PATH', 'CLAUDE_PROXY_TARGET', 'Claude'],
    ['DEEPSEEK_PROXY_PATH', 'DEEPSEEK_PROXY_TARGET', 'DeepSeek'],
    ['GLM_PROXY_PATH', 'GLM_PROXY_TARGET', 'GLM'],
    ['KIMI_PROXY_PATH', 'KIMI_PROXY_TARGET', 'Kimi'],
    ['MINIMAX_PROXY_PATH', 'MINIMAX_PROXY_TARGET', 'MiniMax'],
    ['MINIMAX_CN_PROXY_PATH', 'MINIMAX_CN_PROXY_TARGET', 'MiniMax-CN'],
    ['ZHONGKEYU_PROXY_PATH', 'ZHONGKEYU_PROXY_TARGET', 'ZhongKeYu'],
    ['CUSTOM_PROXY_PATH', 'CUSTOM_PROXY_TARGET', 'Custom'],
  ];

  // webpack-dev-server v5 requires the array form: one entry per provider,
  // with `context` carrying the path prefix (the v4 object map is rejected
  // by its options schema).
  const proxies = [];
  for (const [pathKey, targetKey, label] of providers) {
    const proxyPath = ENV[pathKey];
    const target = ENV[targetKey];
    if (!proxyPath || !target) continue;

    const targetUrl = parseProxyTarget(target);
    if (!targetUrl) {
      console.error(`Ignoring ${targetKey}: target must use HTTPS or loopback HTTP`);
      continue;
    }
    const AgentCtor = targetUrl.protocol === 'https:' ? require('https').Agent : require('http').Agent;

    proxies.push({
      context: [proxyPath],
      target: targetUrl.href,
      changeOrigin: true,
      pathRewrite: { [`^${proxyPath}`]: '' },
      // Verify upstream TLS by default; set LLM_PROXY_TLS_VERIFY=false only
      // for a local LLM backend with a self-signed certificate.
      secure: process.env.LLM_PROXY_TLS_VERIFY !== 'false',
      logLevel: 'debug',
      timeout: LLM_PROXY_TIMEOUT_MS,
      proxyTimeout: LLM_PROXY_TIMEOUT_MS,
      agent: new AgentCtor({
        keepAlive: true,
        keepAliveMsecs: 30000,
        maxSockets: 50,
        maxFreeSockets: 10,
        timeout: LLM_PROXY_TIMEOUT_MS
      }),
      bypass: function (req, res) {
        if (req.method === 'OPTIONS') {
          for (const [name, value] of Object.entries(corsHeaders(req, allowedCorsOrigins))) {
            res.setHeader(name, value);
          }
          res.setHeader('Access-Control-Max-Age', '86400');
          res.statusCode = 204;
          res.end();
          return true;
        }
      },
      onProxyReq: function (proxyReq, req) {
        const requestPath = String(req.url || '').split('?')[0];
        const upstreamPath = String(proxyReq.path || '').split('?')[0];
        console.log(`[${label} Proxy Request]`, req.method, requestPath, '→', upstreamPath);
        if (typeof proxyReq.removeHeader === 'function') {
          proxyReq.removeHeader('origin');
          proxyReq.removeHeader('referer');
        }
      },
      onProxyRes: function (proxyRes, req) {
        const requestPath = String(req.url || '').split('?')[0];
        console.log(`[${label} Proxy Response]`, requestPath, '←', proxyRes.statusCode);
        const corsHeaderNames = new Set([
          'access-control-allow-origin',
          'access-control-allow-methods',
          'access-control-allow-headers',
        ]);
        for (const name of Object.keys(proxyRes.headers)) {
          if (corsHeaderNames.has(name.toLowerCase())) delete proxyRes.headers[name];
        }
        Object.assign(proxyRes.headers, corsHeaders(req, allowedCorsOrigins));
      },
      onError: function (err, req, res) {
        const requestPath = String(req.url || '').split('?')[0];
        console.error(`[${label} Proxy Error]`, requestPath, err.code || 'upstream error');
        if (!res.headersSent) {
          res.writeHead(502, {
            'Content-Type': 'application/json',
            ...corsHeaders(req, allowedCorsOrigins)
          });
          res.end(JSON.stringify({
            error: `${label} Proxy Error`,
            message: 'Upstream request failed',
            code: err.code
          }));
        }
      }
    });
  }
  return proxies;
}

module.exports = (env, argv) => {
  const isDev = argv.mode === 'development';

  // Check if custom certs exist, otherwise use webpack's self-signed cert
  // To use your own certs, place server.pem and server-key.pem in the project root
  // (or set SSL_CERT_FILE and SSL_KEY_FILE environment variables)
  const certPath = process.env.SSL_CERT_FILE
    ? path.resolve(__dirname, process.env.SSL_CERT_FILE)
    : path.resolve(__dirname, 'server.pem');
  const keyPath = process.env.SSL_KEY_FILE
    ? path.resolve(__dirname, process.env.SSL_KEY_FILE)
    : path.resolve(__dirname, 'server-key.pem');

  const customCerts = fs.existsSync(certPath) && fs.existsSync(keyPath);

  // webpack-dev-server v5: TLS is configured via `server`, not `https`.
  // With custom certs → { type: 'https', options: { key, cert } };
  // without → { type: 'https' } generates a self-signed cert at runtime.
  const serverConfig = customCerts ? {
    type: 'https',
    options: {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath)
    }
  } : { type: 'https' };

  return {
    entry: {
      taskpane: './src/taskpane/taskpane.js',
      commands: './src/commands/commands.js'
    },
    output: {
      // Entry names stay stable and hash-free: taskpane.html/commands.html are
      // generated by HtmlWebpackPlugin, but manifest.xml and the sideload
      // scripts reference these URLs directly.
      filename: '[name].js',
      // Async chunks (the mammoth and pdf.js parsers, dynamically imported by
      // file-attachments.js) get semantic names instead of webpack's module
      // ids. Without this the same two ~490 KB parser bundles landed as
      // 255.js/400.js and were renumbered whenever module ids shifted, so a
      // diff of dist/ could not tell a real change from a renumbering.
      chunkFilename: '[name].[contenthash:8].js',
      path: path.resolve(__dirname, 'dist'),
      clean: true
    },
    module: {
      rules: [
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        },
        {
          // The vendored diff-match-patch copy is verbatim CommonJS
          // (module.exports); force CJS parsing so the package.json
          // "type": "module" default doesn't hide its exports.
          test: /src[\\/]lib[\\/]vendor[\\/].+\.js$/,
          type: 'javascript/dynamic'
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: ['commands']
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets',
            to: 'assets',
            noErrorOnMissing: true
          },
          {
            from: 'debug.html',
            to: 'debug.html',
            noErrorOnMissing: true
          },
          {
            // pdf.js worker for .pdf attachments (see src/lib/file-attachments.js).
            // Loaded on demand by pdf.js itself, never by the app bundle.
            //
            // noErrorOnMissing so the worker alone never kills a build: the
            // .pdf attachment path degrades to "unsupported file" at runtime,
            // while a pdfjs-dist layout change on upgrade (or an install
            // without the optional parser) still yields a working add-in.
            // PDF support is one attachment type, not a build dependency.
            from: 'node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs',
            to: 'pdf.worker.min.mjs',
            noErrorOnMissing: true
          }
        ]
      }),
      // Inject environment defaults into the bundle
      new webpack.DefinePlugin({
        'process.env.DEFAULT_OLLAMA_URL': JSON.stringify(ENV.DEFAULT_OLLAMA_URL),
        'process.env.DEFAULT_MODEL': JSON.stringify(ENV.DEFAULT_MODEL),
        'process.env.DEFAULT_VLLM_URL': JSON.stringify(ENV.DEFAULT_VLLM_URL),
        'process.env.DEFAULT_VLLM_MODEL': JSON.stringify(ENV.DEFAULT_VLLM_MODEL),
      })
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, 'dist')
      },
      server: serverConfig,
      host: ENV.DEV_SERVER_HOST,
      port: ENV.DEV_SERVER_PORT,
      hot: true,
      allowedHosts: ENV.DEV_SERVER_ALLOWED_HOSTS,
      headers: (req) => corsHeaders(req, ENV.DEV_SERVER_CORS_ORIGINS),
      setupMiddlewares: process.env.ENABLE_DEV_ENDPOINTS === 'true'
        // Dev-only E2E/coding-agent endpoints (see scripts/dev-e2e-middlewares.cjs).
        // Off by default: they write files and use wildcard CORS, so they are
        // never registered unless explicitly requested.
        ? require('./scripts/dev-e2e-middlewares.cjs')
        : undefined,
      proxy: buildLlmProxies(ENV)
    },
    resolve: {
      extensions: ['.js', '.json']
    },
    performance: {
      // The taskpane bundle legitimately exceeds webpack's 244 KiB default
      // hint: it is served same-origin to WebView2, so there is no cold
      // network fetch and a single ~430 KiB bundle is acceptable. This is a
      // hard gate (hints: 'error' fails the build), calibrated just above
      // the current size, to catch accidental size regressions. Exempt the
      // lazily-loaded parser chunks (mammoth/pdf.js for file attachments)
      // and the pdf.js worker copy: they are fetched on demand, never on
      // first paint.
      hints: 'error',
      maxAssetSize: 700 * 1024,
      maxEntrypointSize: 500 * 1024,
      assetFilter: (name) => !/pdf\.worker\.min\.mjs$/.test(name)
    },
    devtool: isDev ? 'eval-source-map' : false
  };
};

module.exports.__testing = {
  parseAllowedHosts,
  parseCorsOrigins,
  resolveCorsOrigin,
  corsHeaders,
  parseProxyTarget,
  buildLlmProxies,
};

