const fs = require('fs');
const path = require('path');
const http = require('http');
const express = require('express');
const setup = require('../scripts/dev-e2e-middlewares.cjs');
const fixture = require('../scripts/dev-harness-fixture.cjs');
const { bounded, MAX_BYTES, MAX_ENTRIES } = require('../scripts/dev-harness-store.cjs');

const TOKEN = 'test-harness-token-not-a-production-secret';
const paths = ['/log', '/logs', '/logs/clear', '/api/trace-log', '/api/fix-log', '/api/test-cases', '/api/prompts', '/api/prompts/example', ...['status', 'trigger', 'pause', 'claim', 'complete'].map(v => `/api/e2e-loop/${v}`)];
let root;
let servers;
let originalToken;
let originalOrigins;
let consoleSpy;

async function start(app = express()) {
  setup([], { app }, { rootDir: root });
  const server = http.createServer(app);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  servers.push(server);
  return { app, port: server.address().port, origin: `http://127.0.0.1:${server.address().port}` };
}

function request(server, pathname, { method = 'GET', body, raw, token = TOKEN, headers = {} } = {}) {
  const data = raw === undefined ? (body === undefined ? undefined : JSON.stringify(body)) : raw;
  return new Promise((resolve, reject) => {
    const req = http.request({ hostname: '127.0.0.1', port: server.port, path: pathname, method, headers: {
      ...(token === null ? {} : { 'x-claric-harness-token': token }),
      ...(data === undefined ? {} : { 'content-type': 'application/json', 'content-length': Buffer.byteLength(data) }),
      ...headers
    } }, res => {
      let text = '';
      res.setEncoding('utf8');
      res.on('data', chunk => { text += chunk; });
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: text ? JSON.parse(text) : null }));
    });
    req.on('error', reject);
    req.end(data);
  });
}
const post = (server, pathname, body = {}) => request(server, pathname, { method: 'POST', body });
const snapshot = relative => JSON.parse(fs.readFileSync(path.join(root, relative), 'utf8'));

beforeEach(() => {
  root = fs.mkdtempSync(path.join(__dirname, 'dev-harness-tmp-'));
  servers = [];
  originalToken = process.env.CLARIC_HARNESS_TOKEN;
  originalOrigins = process.env.CLARIC_HARNESS_ORIGINS;
  process.env.CLARIC_HARNESS_TOKEN = TOKEN;
  delete process.env.CLARIC_HARNESS_ORIGINS;
  consoleSpy = jest.spyOn(console, 'log').mockImplementation(() => {});
});
afterEach(async () => {
  await Promise.all(servers.map(server => new Promise(resolve => server.close(resolve))));
  jest.restoreAllMocks();
  if (originalToken === undefined) delete process.env.CLARIC_HARNESS_TOKEN;
  else process.env.CLARIC_HARNESS_TOKEN = originalToken;
  if (originalOrigins === undefined) delete process.env.CLARIC_HARNESS_ORIGINS;
  else process.env.CLARIC_HARNESS_ORIGINS = originalOrigins;
  fs.rmSync(root, { recursive: true, force: true });
});

test('every route rejects unauthenticated reads/writes and shares preflight policy', async () => {
  const server = await start();
  for (const route of paths) {
    for (const method of ['GET', 'POST', 'DELETE']) {
      const result = await request(server, route, { method, token: null, headers: { Origin: server.origin } });
      expect(result.status).toBe(401);
      expect(result.headers['access-control-allow-origin']).toBe(server.origin);
    }
    const preflight = await request(server, route, { method: 'OPTIONS', token: null, headers: { Origin: server.origin, 'Access-Control-Request-Method': 'POST', 'Access-Control-Request-Headers': 'Content-Type, X-Claric-Harness-Token' } });
    expect(preflight.status).toBe(204);
    expect(preflight.headers['access-control-allow-headers']).toContain('x-claric-harness-token');
    expect(preflight.headers['access-control-allow-origin']).toBe(server.origin);
  }
  expect((await request(server, '/logs', { token: 'bad' })).status).toBe(401);
  expect((await request(server, '/logs', { token: 'x'.repeat(TOKEN.length) })).status).toBe(401);
  expect((await request(server, '/logs')).status).toBe(200);
});

test('untrusted, null and wildcard origins are denied even with a valid token', async () => {
  const app = express();
  app.use((_req, res, next) => { res.setHeader('Access-Control-Allow-Origin', '*'); next(); });
  const server = await start(app);
  for (const Origin of ['https://hostile.example', 'null', '*', `${server.origin}/`]) {
    for (const method of ['GET', 'POST', 'OPTIONS']) {
      const result = await request(server, '/logs', { method, headers: { Origin } });
      expect(result.status).toBe(403);
      expect(result.headers['access-control-allow-origin']).toBeUndefined();
    }
  }
  expect((await request(server, '/logs', { method: 'OPTIONS', headers: { Origin: server.origin, 'Access-Control-Request-Headers': 'Authorization' } })).status).toBe(403);
  expect((await request(server, '/logs', { method: 'OPTIONS', headers: { 'Access-Control-Request-Method': 'PATCH' } })).status).toBe(403);
});

test('explicit extra origins are allowed but never replace authentication', async () => {
  process.env.CLARIC_HARNESS_ORIGINS = 'https://driver.example';
  const server = await start();
  const headers = { Origin: 'https://driver.example' };
  expect((await request(server, '/logs', { headers })).status).toBe(200);
  expect((await request(server, '/logs', { headers, token: null })).status).toBe(401);
});

test.each(['*', 'null', 'https://driver.example/path'])('invalid origin configuration fails startup: %s', value => {
  process.env.CLARIC_HARNESS_ORIGINS = value;
  expect(() => setup([], { app: express() }, { rootDir: root })).toThrow('exact HTTP(S) origins');
});

test('setup is app-idempotent, generates one token and does not alter process listeners or globals', async () => {
  delete process.env.CLARIC_HARNESS_TOKEN;
  const signals = ['SIGINT', 'SIGTERM', 'exit', 'beforeExit'];
  const listeners = signals.map(signal => process.rawListeners(signal));
  const globals = [global.e2eLogs, global.e2eLoopControl];
  const server = await start();
  const middleware = [];
  expect(setup(middleware, { app: server.app }, { rootDir: root })).toBe(middleware);
  expect(consoleSpy).toHaveBeenCalledTimes(1);
  const token = consoleSpy.mock.calls[0][0].split(': ').pop();
  expect(token).toMatch(/^[0-9a-f-]{36}$/);
  expect((await request(server, '/logs', { token })).status).toBe(200);
  expect(signals.map(signal => process.rawListeners(signal))).toEqual(listeners);
  expect([global.e2eLogs, global.e2eLoopControl]).toEqual(globals);
  expect(fs.readFileSync(path.join(root, 'logs/dev-harness/loop.json'), 'utf8')).not.toContain(token);
});

test.each([null, [], 1, true, 'text'])('non-object body returns 400: %j', async body => {
  const server = await start();
  for (const route of ['/log', '/api/fix-log', '/api/trace-log', '/api/test-cases', '/api/prompts', '/api/e2e-loop/trigger']) {
    expect((await post(server, route, body)).status).toBe(400);
  }
});

test('malformed, unsupported and oversized bodies have controlled errors and CORS', async () => {
  const server = await start();
  for (const [raw, code, extra] of [['{', 400, {}], ['text', 400, { 'content-type': 'text/plain' }], [JSON.stringify({ text: 'x'.repeat(65536) }), 413, {}]]) {
    const result = await request(server, '/log', { method: 'POST', raw, headers: { Origin: server.origin, ...extra } });
    expect(result.status).toBe(code);
    expect(result.headers['access-control-allow-origin']).toBe(server.origin);
  }
});

test('logging, fix and trace writes are immediately persisted with bounded retention', async () => {
  const server = await start();
  fs.writeFileSync(path.join(root, 'logs/e2e-test-logs.json'), JSON.stringify(Array.from({ length: MAX_ENTRIES }, (_, i) => ({ message: `old-${i}` }))));
  expect((await post(server, '/log', { message: 'new', timestamp: '2026-01-01T00:00:00Z' })).body.persisted).toBe(true);
  const entries = snapshot('logs/e2e-test-logs.json').entries;
  expect(entries).toHaveLength(MAX_ENTRIES);
  expect(entries[0].message).toBe('old-1');
  expect(entries.at(-1).message).toBe('new');
  expect((await request(server, '/logs?since=2025-12-31T00:00:00Z')).body).toHaveLength(1);
  expect((await request(server, '/logs?since=invalid')).status).toBe(400);
  expect((await post(server, '/api/fix-log', { message: 'synthetic fix' })).body.totalFixes).toBe(1);
  expect(snapshot('logs/fix-logs.json')).toHaveLength(1);
  expect((await post(server, '/api/trace-log', { testRunNumber: 1, trace: ['synthetic'] })).body.traceLength).toBe(1);
  expect(snapshot('logs/dev-harness/traces.json')).toHaveLength(1);
  expect((await post(server, '/api/trace-log', { testRunNumber: '../../escape', trace: [] })).status).toBe(400);
  expect((await post(server, '/api/trace-log', { testRunNumber: 1 })).status).toBe(400);
  expect((await post(server, '/logs/clear')).body.persisted).toBe(true);
  expect(snapshot('logs/e2e-test-logs.json')).toEqual({ entries: [], lastSequence: 1001 });
  const large = bounded(Array.from({ length: 1000 }, () => ({ text: 'x'.repeat(10000) })));
  expect(Buffer.byteLength(JSON.stringify(large))).toBeLessThanOrEqual(MAX_BYTES);
  expect(large.length).toBeLessThan(1000);
});

test('rename failure returns 500, leaves prior snapshot intact and removes temporary file', async () => {
  const server = await start();
  await post(server, '/log', { message: 'before' });
  const before = snapshot('logs/e2e-test-logs.json');
  jest.spyOn(fs, 'renameSync').mockImplementation(() => { throw new Error('injected disk failure'); });
  for (const route of ['/log', '/logs/clear', '/api/fix-log', '/api/e2e-loop/trigger']) {
    const result = await post(server, route, { message: 'after' });
    expect(result.status).toBe(500);
    expect(result.body).toMatchObject({ success: false, persisted: false });
  }
  expect(snapshot('logs/e2e-test-logs.json')).toEqual(before);
  expect(fs.readdirSync(path.join(root, 'logs')).some(name => name.endsWith('.tmp'))).toBe(false);
  expect((await request(server, '/api/e2e-loop/status')).body.state).toBe('paused');
});

test('corrupt and oversized snapshots fail closed instead of being silently erased', async () => {
  const server = await start();
  for (const data of ['{', '{}', 'x'.repeat(MAX_BYTES + 1)]) {
    fs.writeFileSync(path.join(root, 'logs/e2e-test-logs.json'), data);
    expect((await post(server, '/log', { message: 'new' })).status).toBe(500);
    expect(fs.readFileSync(path.join(root, 'logs/e2e-test-logs.json'), 'utf8')).toBe(data);
  }
});

test('symlink storage cannot redirect writes', async () => {
  const server = await start();
  const target = path.join(root, 'untouched.json');
  fs.writeFileSync(target, '[]');
  fs.symlinkSync(target, path.join(root, 'logs/e2e-test-logs.json'));
  expect((await post(server, '/log', { message: 'no' })).status).toBe(500);
  expect(fs.readFileSync(target, 'utf8')).toBe('[]');
});

test('loop starts paused, allows exactly one claim and requires current live run for completion', async () => {
  const server = await start();
  expect((await request(server, '/api/e2e-loop/status')).body).toMatchObject({ canProceed: false, waitingForTrigger: true, runId: null, state: 'paused' });
  expect((await post(server, '/api/e2e-loop/claim', { runId: 'absent' })).status).toBe(409);
  const run = (await post(server, '/api/e2e-loop/trigger')).body;
  expect(run.canProceed).toBe(true);
  expect((await post(server, '/api/e2e-loop/trigger')).status).toBe(409);
  expect((await post(server, '/api/e2e-loop/complete', { runId: run.runId, outcome: 'passed' })).status).toBe(409);
  const claims = await Promise.all([post(server, '/api/e2e-loop/claim', { runId: run.runId }), post(server, '/api/e2e-loop/claim', { runId: run.runId })]);
  expect(claims.map(result => result.status).sort()).toEqual([200, 409]);
  expect((await request(server, '/api/e2e-loop/status')).body.canProceed).toBe(false);
  const complete = await post(server, '/api/e2e-loop/complete', { runId: run.runId, outcome: 'passed' });
  expect(complete.body.lastIteration).toMatchObject({ runId: run.runId, outcome: 'passed' });
  expect(snapshot('logs/dev-harness/loop.json').run.state).toBe('completed');
  expect((await post(server, '/api/e2e-loop/complete', { runId: run.runId, outcome: 'passed' })).status).toBe(409);
  const next = (await post(server, '/api/e2e-loop/trigger')).body;
  expect((await post(server, '/api/e2e-loop/claim', { runId: run.runId })).status).toBe(409);
  await post(server, '/api/e2e-loop/pause', { runId: next.runId });
  expect((await post(server, '/api/e2e-loop/claim', { runId: next.runId })).status).toBe(409);
});

test.each([false, true])('unclaimed or claimed runs expire at their deadline (claimed=%s)', async claimed => {
  const server = await start();
  const run = (await post(server, '/api/e2e-loop/trigger')).body;
  if (claimed) await post(server, '/api/e2e-loop/claim', { runId: run.runId });
  jest.spyOn(Date, 'now').mockReturnValue(Date.parse(run.expiresAt));
  expect((await post(server, '/api/e2e-loop/claim', { runId: run.runId })).status).toBe(410);
  expect((await post(server, '/api/e2e-loop/complete', { runId: run.runId, outcome: 'passed' })).status).toBe(410);
  expect((await post(server, '/api/e2e-loop/pause', { runId: run.runId })).status).toBe(410);
  expect((await request(server, '/api/e2e-loop/status')).body.state).toBe('expired');
  expect((await post(server, '/api/e2e-loop/trigger')).status).toBe(200);
});

test('a new app instance resets paused and rejects the prior runId', async () => {
  const first = await start();
  const run = (await post(first, '/api/e2e-loop/trigger')).body;
  await post(first, '/api/e2e-loop/claim', { runId: run.runId });
  const second = await start();
  expect((await request(second, '/api/e2e-loop/status')).body.state).toBe('paused');
  expect((await post(second, '/api/e2e-loop/claim', { runId: run.runId })).status).toBe(409);
});

test('synthetic plain-text fixture and prompt CRUD traverse real HTTP and isolated storage', async () => {
  const server = await start();
  expect(fixture.original.trim()).not.toBe('');
  expect(fixture.modified.trim()).not.toBe('');
  const added = await post(server, '/api/test-cases', fixture);
  expect(added.status).toBe(200);
  expect(added.body.testCase.expected).toBe(fixture.modified);
  expect((await request(server, '/api/test-cases')).body).toEqual([added.body.testCase]);
  expect(fs.existsSync(path.join(root, 'e2e-test-cases-dynamic.json'))).toBe(false);
  const defaults = [{ id: 'example', name: 'Default', template: '{selection}' }];
  fs.writeFileSync(path.join(root, 'prompts.json'), JSON.stringify(defaults));
  const prompt = { id: 'example', name: 'Synthetic', template: 'Edit {selection}' };
  expect((await post(server, '/api/prompts', prompt)).body.prompt).toEqual(prompt);
  expect((await request(server, '/api/prompts')).body).toEqual([prompt]);
  expect((await request(server, '/api/prompts/example', { method: 'DELETE' })).status).toBe(200);
  expect((await request(server, '/api/prompts')).body).toEqual(defaults);
  expect((await request(server, '/api/prompts/example', { method: 'DELETE' })).status).toBe(404);
  expect((await post(server, '/api/prompts', { id: 4, name: true, template: [] })).status).toBe(400);
  expect((await post(server, '/api/test-cases', { original: [], modified: 'text' })).status).toBe(400);
});

test('legacy custom files migrate on first write without modifying original files', async () => {
  fs.writeFileSync(path.join(root, 'e2e-test-cases-dynamic.json'), JSON.stringify([fixture]));
  const legacyPrompt = { id: 'legacy', name: 'Legacy', template: '{selection}' };
  fs.writeFileSync(path.join(root, 'user-prompts.json'), JSON.stringify([legacyPrompt]));
  const server = await start();
  expect((await request(server, '/api/test-cases')).body).toEqual([fixture]);
  expect((await request(server, '/api/prompts')).body).toEqual([legacyPrompt]);
  expect((await post(server, '/api/test-cases', { ...fixture, id: 'second' })).status).toBe(200);
  expect(snapshot('logs/dev-harness/test-cases.json')).toHaveLength(2);
  expect(snapshot('e2e-test-cases-dynamic.json')).toEqual([fixture]);
  expect((await request(server, '/api/prompts/legacy', { method: 'DELETE' })).status).toBe(200);
  expect((await request(server, '/api/prompts')).body).toEqual([]);
  expect(snapshot('user-prompts.json')).toEqual([legacyPrompt]);
});

test('disk write failure cannot acknowledge logs, traces, custom records or loop changes', async () => {
  const server = await start();
  await post(server, '/api/prompts', { id: 'example', name: 'Example', template: '{selection}' });
  jest.spyOn(fs, 'writeFileSync').mockImplementation(() => { throw new Error('injected ENOSPC'); });
  for (const [route, body] of [
    ['/log', { message: 'synthetic' }],
    ['/api/fix-log', { message: 'synthetic' }],
    ['/api/trace-log', { testRunNumber: 1, trace: [] }],
    ['/api/test-cases', fixture],
    ['/api/prompts', { id: 'another', name: 'Another', template: '{selection}' }],
    ['/api/e2e-loop/pause', {}]
  ]) {
    const result = await post(server, route, body);
    expect(result.status).toBe(500);
    expect(result.body.persisted).toBe(false);
  }
  expect((await request(server, '/api/prompts/example', { method: 'DELETE' })).status).toBe(500);
  expect((await request(server, '/api/prompts')).body).toHaveLength(1);
});

test('invalid tokens fail startup and original non-harness routes remain unaffected', async () => {
  process.env.CLARIC_HARNESS_TOKEN = 'short';
  expect(() => setup([], { app: express() }, { rootDir: root })).toThrow('32-256');
  process.env.CLARIC_HARNESS_TOKEN = TOKEN;
  const server = await start();
  server.app.get('/unrelated', (_req, res) => res.json({ untouched: true }));
  expect((await request(server, '/unrelated', { token: null })).body).toEqual({ untouched: true });
  expect((await request(server, '/logs', { method: 'PATCH' })).status).toBe(405);
});
