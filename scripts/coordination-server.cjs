const http = require('http');
const https = require('https');
const fs = require('fs');
const net = require('net');
const crypto = require('crypto');
const { TextDecoder } = require('util');
const { CoordinationStore } = require('../src/lib/coordination/store.cjs');
const { CoordinationState } = require('../src/lib/coordination/state.cjs');
const { normalizeId } = require('../src/lib/coordination/identity.cjs');

const MAX_BODY_BYTES = 128 * 1024;
function isLoopback(address) {
  if (address === '::1') return true;
  const ipv4 = String(address || '').replace(/^::ffff:/i, '');
  return net.isIP(ipv4) === 4 && ipv4.startsWith('127.');
}
function json(res, status, body) {
  if (res.writableEnded || res.destroyed || res.headersSent) return;
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  res.end(JSON.stringify(body));
}
function errorWithStatus(message, status = 400) { return Object.assign(new Error(message), { status }); }
function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    let settled = false;
    const chunks = [];
    const fail = (error) => { if (!settled) { settled = true; chunks.length = 0; reject(error); } };
    req.on('data', (chunk) => {
      if (settled) return;
      size += chunk.length;
      if (size > MAX_BODY_BYTES) { fail(errorWithStatus('Request body too large', 413)); return; }
      chunks.push(chunk);
    });
    req.on('end', () => {
      if (settled) return;
      settled = true;
      try {
        const text = new TextDecoder('utf-8', { fatal: true }).decode(Buffer.concat(chunks));
        resolve(JSON.parse(text));
      } catch { reject(errorWithStatus('Invalid JSON')); }
    });
    req.on('aborted', () => fail(errorWithStatus('Request aborted')));
    req.on('error', () => fail(errorWithStatus('Request body unavailable')));
    req.on('close', () => { if (!req.complete) fail(errorWithStatus('Incomplete request')); });
    if (Number(req.headers['content-length']) > MAX_BODY_BYTES) fail(errorWithStatus('Request body too large', 413));
  });
}
function authorized(req, token) {
  if (!token) return true;
  const digest = (value) => crypto.createHash('sha256').update(value).digest();
  return crypto.timingSafeEqual(digest(typeof req.headers.authorization === 'string' ? req.headers.authorization : ''), digest(`Bearer ${token}`));
}
function createCoordinationHandler({ store = new CoordinationStore(), state, persistencePath, token = process.env.COORDINATION_TOKEN || '', allowedOrigin = process.env.COORDINATION_ALLOWED_ORIGIN || '' } = {}) {
  if (!state) state = new CoordinationState(persistencePath ? { persistencePath } : {});
  if (allowedOrigin && (!/^https?:\/\//.test(allowedOrigin) || new URL(allowedOrigin).origin !== allowedOrigin)) throw new Error('Configure one exact HTTP(S) coordination origin');
  return async function coordinationHandler(req, res) {
    if (!isLoopback(req.socket?.remoteAddress)) return json(res, 403, { error: 'Loopback only' });
    let host;
    let url;
    try {
      host = new URL(`http://${req.headers.host}`);
      if (host.username || host.password || host.pathname !== '/' || host.search || host.hash || !(host.hostname === 'localhost' || host.hostname === '[::1]' || isLoopback(host.hostname))) throw new Error();
      if (!req.url.startsWith('/') || req.url.startsWith('//') || /[\\#]/.test(req.url) || /%(?![\da-f]{2})/i.test(req.url)) throw new Error();
      url = new URL(req.url, host);
    } catch { return json(res, 400, { error: 'Invalid Host or request URL' }); }
    const origin = req.headers.origin;
    const sameOrigin = `${req.socket.encrypted ? 'https:' : 'http:'}//${host.host}`;
    if (origin && origin !== sameOrigin && origin !== allowedOrigin) return json(res, 403, { error: 'Origin not allowed' });
    if (origin) { res.setHeader('Access-Control-Allow-Origin', origin); res.setHeader('Vary', 'Origin'); }
    const v2 = url.pathname.startsWith('/coordination/v2/');
    const registering = url.pathname === '/coordination/v2/instances/register';
    const method = {
      '/coordination/healthz': 'GET',
      '/coordination/snapshot': 'GET',
      '/coordination/events': 'POST',
      '/coordination/v2/instances/register': 'POST',
      '/coordination/v2/snapshot': 'GET',
      '/coordination/v2/events': 'GET',
      '/coordination/v2/envelopes': 'POST',
    }[url.pathname];
    if (!method) return json(res, 404, { error: 'Not found' });
    const params = [...url.searchParams.keys()];
    const expectedParams = url.pathname === '/coordination/snapshot' ? ['workspaceId']
      : url.pathname === '/coordination/v2/events' ? ['after', 'epoch'] : [];
    if (params.length !== expectedParams.length || expectedParams.some((name) => url.searchParams.getAll(name).length !== 1)) return json(res, 400, { error: 'Invalid query parameters' });
    if (req.method === 'OPTIONS') {
      const requestedHeaders = (req.headers['access-control-request-headers'] || '').toLowerCase().split(',').map((value) => value.trim()).filter(Boolean);
      if (!origin || req.headers['access-control-request-method'] !== method || requestedHeaders.some((name) => !['authorization', 'content-type'].includes(name))) return json(res, 403, { error: 'Invalid preflight' });
      res.writeHead(204, { 'Access-Control-Allow-Methods': method, 'Access-Control-Allow-Headers': 'Authorization, Content-Type', 'Access-Control-Max-Age': '600' });
      return res.end();
    }
    const credential = typeof req.headers.authorization === 'string' && req.headers.authorization.startsWith('Bearer ')
      ? req.headers.authorization.slice(7) : '';
    const binding = JSON.stringify([sameOrigin, origin || null]);
    if (registering) {
      if (token && !authorized(req, token)) return json(res, 401, { error: 'Unauthorized' });
    } else if (!v2) {
      if (!authorized(req, token)) return json(res, 401, { error: 'Unauthorized' });
    } else {
      try { state.authenticate(credential, binding); }
      catch (error) { return json(res, error.status || 401, { error: error.message }); }
    }
    if (req.method !== method) { res.setHeader('Allow', method); return json(res, 405, { error: 'Method not allowed' }); }
    if (url.pathname === '/coordination/healthz') return json(res, 200, { status: 'ok', service: 'coordination', version: 2 });
    try {
      if (url.pathname === '/coordination/v2/snapshot') return json(res, 200, state.snapshot(credential, binding));
      if (url.pathname === '/coordination/v2/events') {
        const after = url.searchParams.get('after');
        if (!/^(0|[1-9][0-9]*)$/.test(after)) return json(res, 400, { error: 'Invalid cursor' });
        return json(res, 200, state.events(credential, binding, Number(after), url.searchParams.get('epoch')));
      }
      if (url.pathname === '/coordination/snapshot') return json(res, 200, store.snapshot(normalizeId(url.searchParams.get('workspaceId'), 'workspaceId')));
      if (registering || req.method === 'POST') {
        if (!/^application\/json(?:\s*;\s*charset\s*=\s*(?:utf-8|"utf-8"))?\s*$/i.test(req.headers['content-type'] || '') || req.headers['content-encoding']) return json(res, 415, { error: 'Use uncompressed application/json UTF-8' });
      }
      const body = req.method === 'POST' ? await readBody(req) : null;
      if (registering) return json(res, 201, { version: 2, ...state.register(body, binding) });
      if (v2) return json(res, 201, { event: state.append(body, credential, binding) });
      return json(res, 201, { event: store.append(body) });
    } catch (error) { return json(res, error.status || 400, { error: error.message || 'Invalid request' }); }
  };
}
function createCoordinationServer(options = {}) {
  const handler = createCoordinationHandler(options);
  const listener = (req, res) => { handler(req, res).catch(() => json(res, 500, { error: 'Internal server error' })); };
  const server = options.tls ? https.createServer(options.tls, listener) : http.createServer(listener);
  server.requestTimeout = 15000;
  server.headersTimeout = 10000;
  const listen = server.listen.bind(server);
  server.listen = (port, host = '127.0.0.1', callback) => {
    if (typeof host === 'function') { callback = host; host = '127.0.0.1'; }
    if (host === 'localhost') host = '127.0.0.1';
    if (!isLoopback(host) || !Number.isInteger(Number(port)) || Number(port) < 0 || Number(port) > 65535) throw new Error('Use a loopback host and valid TCP port');
    return listen(Number(port), host, callback);
  };
  return server;
}
if (require.main === module) {
  const host = process.env.COORDINATION_HOST || '127.0.0.1';
  const port = Number(process.env.COORDINATION_PORT || 3010);
  const cert = process.env.COORDINATION_CERT_FILE;
  const key = process.env.COORDINATION_KEY_FILE;
  if (!!cert !== !!key) throw new Error('Provide both coordination TLS certificate and key');
  const persistencePath = process.env.COORDINATION_STATE_FILE || '';
  const server = createCoordinationServer({ tls: cert ? { cert: fs.readFileSync(cert), key: fs.readFileSync(key) } : undefined, persistencePath: persistencePath || undefined });
  server.on('error', (error) => { console.error(`Coordination server: ${error.code || 'startup failed'}`); process.exitCode = 1; });
  server.listen(port, host, () => console.log(`Coordination server listening on loopback port ${port} (${cert ? 'HTTPS' : 'HTTP'})`));
}
module.exports = { isLoopback, authorized, readBody, createCoordinationHandler, createCoordinationServer, MAX_BODY_BYTES };
