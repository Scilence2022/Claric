const path = require('path');
const { randomUUID, timingSafeEqual } = require('crypto');
const { createStore, bounded, MAX_ENTRIES } = require('./dev-harness-store.cjs');

const instances = new WeakSet();
const TOKEN_HEADER = 'x-claric-harness-token';
const BODY_LIMIT = 64 * 1024;
const RUN_TTL_MS = 5 * 60 * 1000;
const METHODS = 'GET, POST, DELETE, OPTIONS';
const ROUTES = /^\/(?:log|logs(?:\/clear)?|api\/(?:trace-log|fix-log|test-cases|prompts(?:\/[^/]+)?|e2e-loop\/(?:status|trigger|pause|claim|complete)))\/?$/i;

function origin(value) {
  try {
    const url = new URL(value);
    return /^https?:$/.test(url.protocol) && url.origin === value ? value : null;
  } catch (_error) { return null; }
}

function fail(status, message) {
  const error = new Error(message);
  error.status = status;
  throw error;
}

function nonempty(value) { return typeof value === 'string' && value.trim().length > 0; }

function setupDevE2eMiddlewares(middlewares, devServer, options = {}) {
  if (!devServer?.app) throw new Error('webpack-dev-server is not defined');
  const app = devServer.app;
  if (instances.has(app)) return middlewares;
  if (options.rootDir && process.env.NODE_ENV !== 'test') throw new Error('rootDir is test-only');
  const rootDir = options.rootDir || path.resolve(__dirname, '..');
  const store = createStore(rootDir);
  const token = process.env.CLARIC_HARNESS_TOKEN || randomUUID();
  if (!/^[\x21-\x7e]{32,256}$/.test(token)) throw new Error('CLARIC_HARNESS_TOKEN must be 32-256 printable non-space ASCII characters');
  const extraOrigins = (process.env.CLARIC_HARNESS_ORIGINS || '').split(',').map(v => v.trim()).filter(Boolean);
  if (extraOrigins.some(value => !origin(value))) throw new Error('CLARIC_HARNESS_ORIGINS requires exact HTTP(S) origins, never *');
  const tokenBytes = Buffer.from(token);
  const express = require('express');
  const router = express.Router();
  let run = null;
  let lastIteration = null;
  const fixesPath = 'logs/fix-logs.json';
  const dynamicPath = 'logs/dev-harness/test-cases.json';
  const promptsPath = 'logs/dev-harness/user-prompts.json';
  const runPath = 'logs/dev-harness/loop.json';

  function saveRun(next, iteration = lastIteration) {
    store.write(runPath, { run: next, lastIteration: iteration });
    run = next;
    lastIteration = iteration;
  }
  function expire() {
    if (run && ['pending', 'claimed'].includes(run.state) && Date.now() >= run.expiresAt) {
      saveRun({ ...run, state: 'expired' });
    }
  }
  function status() {
    expire();
    return {
      canProceed: run?.state === 'pending',
      waitingForTrigger: run?.state !== 'pending',
      lastIteration,
      runId: run?.runId || null,
      state: run?.state || 'paused',
      expiresAt: run ? new Date(run.expiresAt).toISOString() : null
    };
  }
  function requireRun(body, state) {
    if (!nonempty(body.runId)) fail(400, 'runId is required');
    expire();
    if (!run || run.runId !== body.runId) fail(409, 'Unknown or superseded runId');
    if (run.state === 'expired') fail(410, 'Run expired');
    if (run.state !== state) fail(409, `Run is ${run.state}`);
  }
  function append(relative, entry) {
    const entries = bounded([...store.readArray(relative), { ...entry, receivedAt: new Date().toISOString() }]);
    store.write(relative, entries);
    return entries.length;
  }
  function custom(relative, legacy) {
    const value = store.read(relative, null);
    if (value !== null) {
      if (!Array.isArray(value)) throw new Error('Stored snapshot must be an array');
      return value;
    }
    return store.readArray(legacy);
  }

  router.use((req, res, next) => {
    for (const header of ['Access-Control-Allow-Origin', 'Access-Control-Allow-Credentials', 'Access-Control-Allow-Methods', 'Access-Control-Allow-Headers']) res.removeHeader(header);
    res.vary('Origin');
    res.setHeader('Cache-Control', 'no-store');
    const requestOrigin = req.get('Origin');
    const protocol = req.socket.encrypted ? 'https' : 'http';
    const serviceOrigin = `${protocol}://${req.get('Host')}`;
    if (requestOrigin && (!origin(requestOrigin) || (requestOrigin !== serviceOrigin && !extraOrigins.includes(requestOrigin)))) {
      return res.status(403).json({ success: false, error: 'Origin is not allowed' });
    }
    if (requestOrigin) res.setHeader('Access-Control-Allow-Origin', requestOrigin);
    res.setHeader('Access-Control-Allow-Methods', METHODS);
    res.setHeader('Access-Control-Allow-Headers', `Content-Type, ${TOKEN_HEADER}`);
    if (req.method === 'OPTIONS') {
      const method = req.get('Access-Control-Request-Method');
      const headers = (req.get('Access-Control-Request-Headers') || '').split(',').map(v => v.trim().toLowerCase()).filter(Boolean);
      if ((method && !METHODS.split(', ').includes(method)) || headers.some(v => !['content-type', TOKEN_HEADER].includes(v))) {
        return res.status(403).json({ success: false, error: 'Preflight is not allowed' });
      }
      return res.status(204).end();
    }
    const supplied = Buffer.from(req.get(TOKEN_HEADER) || '');
    if (supplied.length !== tokenBytes.length || !timingSafeEqual(supplied, tokenBytes)) {
      return res.status(401).json({ success: false, error: 'Harness token required' });
    }
    next();
  });
  router.use((req, _res, next) => {
    const hasBody = req.headers['transfer-encoding'] || Number(req.headers['content-length']) > 0;
    if (Number(req.headers['content-length']) > BODY_LIMIT) fail(413, 'JSON body exceeds 64 KiB');
    if (hasBody && !req.is('application/json')) fail(400, 'Content-Type must be application/json');
    next();
  });
  router.use(express.json({ limit: BODY_LIMIT, strict: false, inflate: false }));
  router.use((req, _res, next) => {
    const hasBody = req.headers['transfer-encoding'] || Number(req.headers['content-length']) > 0;
    if (hasBody && (!req.body || typeof req.body !== 'object' || Array.isArray(req.body))) fail(400, 'JSON body must be an object');
    if (!hasBody) req.body = {};
    next();
  });

  const logs = require('./dev-harness-logs.cjs').createLogs(store);
  router.post('/log', (req, res) => {
    const seq = logs.append(req.body);
    res.json({ success: true, persisted: true, seq });
  });
  router.get('/logs', (req, res) => {
    if (req.query.after !== undefined) {
      if (req.query.since !== undefined || typeof req.query.after !== 'string' || !/^(0|[1-9]\d*)$/.test(req.query.after) || !Number.isSafeInteger(Number(req.query.after))) fail(400, 'after must be a nonnegative safe integer and cannot be combined with since');
      return res.json(logs.after(Number(req.query.after)));
    }
    const entries = logs.read().entries;
    const since = req.query.since ? Date.parse(req.query.since) : null;
    if (since !== null && !Number.isFinite(since)) fail(400, 'Invalid since timestamp');
    res.json(since === null ? entries : entries.filter(entry => Date.parse(entry.timestamp || entry.receivedAt) > since));
  });
  router.post('/logs/clear', (_req, res) => {
    logs.clear();
    res.json({ success: true, persisted: true, message: 'Logs cleared from memory and file' });
  });
  router.post('/api/fix-log', (req, res) => {
    const totalFixes = append(fixesPath, req.body);
    res.json({ success: true, persisted: true, message: 'Fix logged successfully', totalFixes });
  });
  router.post('/api/trace-log', (req, res) => {
    const data = req.body;
    if (!/^\d{1,15}$/.test(String(data.testRunNumber)) || !Number.isSafeInteger(Number(data.testRunNumber))) fail(400, 'testRunNumber must be a nonnegative safe integer of at most 15 digits');
    if (!Array.isArray(data.trace)) fail(400, 'trace must be an array');
    const traceFileName = `trace-log-${Number(data.testRunNumber)}.json`;
    const traces = bounded([...store.readArray('logs/dev-harness/traces.json'), { ...data, receivedAt: new Date().toISOString() }]);
    store.write('logs/dev-harness/traces.json', traces);
    res.json({ success: true, persisted: true, message: `Trace saved to bounded snapshot (${traceFileName})`, traceLength: data.trace.length });
  });

  router.get('/api/e2e-loop/status', (_req, res) => res.json(status()));
  router.post('/api/e2e-loop/trigger', (_req, res) => {
    expire();
    if (run && ['pending', 'claimed'].includes(run.state)) fail(409, 'An active run already exists');
    saveRun({ runId: randomUUID(), state: 'pending', expiresAt: Date.now() + RUN_TTL_MS });
    res.json({ success: true, persisted: true, message: 'Loop trigger activated', ...status() });
  });
  router.post('/api/e2e-loop/claim', (req, res) => {
    requireRun(req.body, 'pending');
    saveRun({ ...run, state: 'claimed' });
    res.json({ success: true, persisted: true, ...status() });
  });
  router.post('/api/e2e-loop/pause', (req, res) => {
    if (req.body.runId !== undefined) {
      if (!nonempty(req.body.runId)) fail(400, 'runId must be a nonempty string');
      expire();
      if (!run || run.runId !== req.body.runId) fail(409, 'Unknown or superseded runId');
      if (run.state === 'expired') fail(410, 'Run expired');
    }
    saveRun(run ? { ...run, state: 'paused' } : null);
    res.json({ success: true, persisted: true, message: 'Loop paused, waiting for trigger', ...status() });
  });
  router.post('/api/e2e-loop/complete', (req, res) => {
    requireRun(req.body, 'claimed');
    if (!['passed', 'failed'].includes(req.body.outcome)) fail(400, 'outcome must be passed or failed');
    const iteration = { runId: run.runId, outcome: req.body.outcome, completedAt: new Date().toISOString() };
    saveRun({ ...run, state: 'completed' }, iteration);
    res.json({ success: true, persisted: true, ...status() });
  });

  router.get('/api/test-cases', (_req, res) => {
    res.json([...store.readArray('e2e-test-cases.json'), ...custom(dynamicPath, 'e2e-test-cases-dynamic.json')]);
  });
  router.post('/api/test-cases', (req, res) => {
    const body = req.body;
    if (!nonempty(body.original) || !nonempty(body.modified)) fail(400, 'Invalid test case: original and modified must be nonempty strings');
    if (body.id !== undefined && !nonempty(body.id)) fail(400, 'id must be a nonempty string');
    const entries = custom(dynamicPath, 'e2e-test-cases-dynamic.json');
    if (entries.length >= MAX_ENTRIES) fail(413, 'Test case capacity exceeded');
    const testCase = { id: body.id || `test-${randomUUID()}`, original: body.original, modified: body.modified, expected: body.modified, reason: body.reason || 'auto-generated', createdAt: new Date().toISOString() };
    entries.push(testCase);
    store.write(dynamicPath, entries);
    res.json({ success: true, persisted: true, testCase });
  });
  router.get('/api/prompts', (_req, res) => {
    const entries = [...store.readArray('prompts.json'), ...custom(promptsPath, 'user-prompts.json')];
    res.json(Array.from(new Map(entries.map(entry => [entry.id, entry])).values()));
  });
  router.post('/api/prompts', (req, res) => {
    const body = req.body;
    if (!['id', 'name', 'template'].every(key => nonempty(body[key]))) fail(400, 'Invalid prompt: id, name and template must be nonempty strings');
    const entries = custom(promptsPath, 'user-prompts.json');
    const index = entries.findIndex(entry => entry.id === body.id);
    if (index >= 0) entries[index] = body;
    else {
      if (entries.length >= MAX_ENTRIES) fail(413, 'Prompt capacity exceeded');
      entries.push(body);
    }
    store.write(promptsPath, entries);
    res.json({ success: true, persisted: true, prompt: body });
  });
  router.delete('/api/prompts/:id', (req, res) => {
    const entries = custom(promptsPath, 'user-prompts.json');
    const filtered = entries.filter(entry => entry.id !== req.params.id);
    if (filtered.length === entries.length) fail(404, 'Prompt not found or cannot delete default prompt');
    store.write(promptsPath, filtered);
    res.json({ success: true, persisted: true, message: `Deleted prompt: ${req.params.id}` });
  });
  router.use((_req, res) => res.status(405).json({ success: false, error: 'Method not allowed' }));
  router.use((error, _req, res, _next) => {
    const code = error.status === 413 ? 413 : (error.status >= 400 && error.status < 500 ? error.status : 500);
    res.status(code).json({ success: false, persisted: false, error: code === 500 ? 'Harness storage operation failed' : (error.type ? 'Invalid or oversized JSON body' : error.message) });
  });

  // Persist the restart boundary before registering routes; never resume a previous claim.
  saveRun(null, null);
  app.use((req, res, next) => ROUTES.test(req.path) ? router(req, res, next) : next());
  instances.add(app);
  console.log(`[Claric Harness] Local driver token (${TOKEN_HEADER}): ${token}`);
  return middlewares;
}

module.exports = setupDevE2eMiddlewares;
