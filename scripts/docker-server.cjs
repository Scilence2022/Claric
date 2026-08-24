/**
 * Production static file server for the Word AI Redliner add-in.
 *
 * Serves the webpack build output (dist/) plus the generated manifest.xml
 * over HTTPS (default) or HTTP (PROTOCOL=http) for local testing.
 *
 * Environment variables:
 *   PORT            - listen port (default 3000)
 *   PROTOCOL        - 'https' (default) or 'http'
 *   SSL_CERT_FILE   - cert path, relative to project root or absolute
 *   SSL_KEY_FILE    - key path, relative to project root or absolute
 *
 * Note: this server intentionally does NOT proxy LLM traffic (unlike the
 * webpack dev server). In production the add-in talks to the LLM backend
 * directly, so its endpoint URL must be absolute and reachable from the
 * client machine.
 */

const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');
const { generateManifest } = require('./generate-manifest.cjs');

const rootDir = path.resolve(__dirname, '..');
const distDir = path.join(rootDir, 'dist');
const manifestPath = path.join(rootDir, 'manifest.xml');

// Requests are held open by Word until the page loads; give slow clients
// a bounded window instead of the 2-minute node default.
const REQUEST_TIMEOUT_MS = 60000;
const SHUTDOWN_TIMEOUT_MS = 10000;

function getEnv() {
  return {
    PORT: process.env.PORT || '3000',
    PROTOCOL: process.env.PROTOCOL || 'https',
    SSL_CERT_FILE: process.env.SSL_CERT_FILE || 'server.pem',
    SSL_KEY_FILE: process.env.SSL_KEY_FILE || 'server-key.pem'
  };
}

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

function requestHandler(req, res) {
  let urlPath;
  try {
    urlPath = decodeURIComponent(req.url.split('?')[0]);
  } catch {
    // Malformed percent-encoding (e.g. GET /%) -- reject, never crash.
    sendError(res, 400, 'Bad request');
    return;
  }

  if (req.method !== 'GET' && req.method !== 'HEAD') {
    res.setHeader('Allow', 'GET, HEAD');
    sendError(res, 405, 'Method not allowed');
    return;
  }

  res.on('finish', () => {
    console.log(`${new Date().toISOString()} ${req.method} ${req.url} -> ${res.statusCode}`);
  });

  if (urlPath === '/healthz') {
    res.writeHead(200, {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    });
    res.end(JSON.stringify({ status: 'ok' }));
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

startServer();
