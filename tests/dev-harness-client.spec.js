const fs = require('fs');
const path = require('path');
const http = require('http');
const { spawn } = require('child_process');
const express = require('express');
const setup = require('../scripts/dev-e2e-middlewares.cjs');
const { baseUrl, command } = require('../scripts/dev-harness-client.cjs');

const TOKEN = 'synthetic-client-token-not-a-real-secret';
let root;
let servers;
let url;
let savedToken;
let savedOrigins;

async function start(app = express()) {
  setup([], { app }, { rootDir: root });
  const server = http.createServer(app);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  servers.push(server);
  return `http://127.0.0.1:${server.address().port}`;
}
async function request(endpoint, body) {
  const response = await fetch(`${url}${endpoint}`, { method: body === undefined ? 'GET' : 'POST', headers: { 'Content-Type': 'application/json', 'x-claric-harness-token': TOKEN }, ...(body === undefined ? {} : { body: JSON.stringify(body) }) });
  return { status: response.status, body: await response.json() };
}
function cli(args, token = TOKEN) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [path.join(__dirname, '../scripts/dev-harness-client.cjs'), ...args, '--base-url', url], { env: { ...process.env, CLARIC_HARNESS_TOKEN: token } });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', data => { stdout += data; });
    child.stderr.on('data', data => { stderr += data; });
    child.on('error', reject);
    child.on('close', code => resolve({ code, stdout, stderr, body: stdout ? JSON.parse(stdout) : null }));
  });
}

beforeEach(async () => {
  root = fs.mkdtempSync(path.join(__dirname, 'dev-harness-client-tmp-'));
  servers = [];
  savedToken = process.env.CLARIC_HARNESS_TOKEN;
  savedOrigins = process.env.CLARIC_HARNESS_ORIGINS;
  process.env.CLARIC_HARNESS_TOKEN = TOKEN;
  delete process.env.CLARIC_HARNESS_ORIGINS;
  jest.spyOn(console, 'log').mockImplementation(() => {});
  url = await start();
});
afterEach(async () => {
  await Promise.all(servers.map(server => new Promise(resolve => { server.closeAllConnections(); server.close(resolve); })));
  jest.restoreAllMocks();
  if (savedToken === undefined) delete process.env.CLARIC_HARNESS_TOKEN;
  else process.env.CLARIC_HARNESS_TOKEN = savedToken;
  if (savedOrigins === undefined) delete process.env.CLARIC_HARNESS_ORIGINS;
  else process.env.CLARIC_HARNESS_ORIGINS = savedOrigins;
  fs.rmSync(root, { recursive: true, force: true });
});

test('sequence polling reads every timestamp-free, identical-time and malformed-time log', async () => {
  expect((await request('/logs?after=0')).body).toEqual({ entries: [], nextCursor: 0, oldestCursor: 1, gap: false });
  for (const body of [{ message: 'no timestamp', seq: 999 }, { message: 'bad timestamp', timestamp: 'invalid' }, { message: 'one', timestamp: '2020-01-01' }, { message: 'two', timestamp: '2020-01-01' }]) {
    expect((await request('/log', body)).status).toBe(200);
  }
  const first = (await request('/logs?after=0')).body;
  expect(first.entries.map(entry => entry.seq)).toEqual([1, 2, 3, 4]);
  expect(first).toMatchObject({ nextCursor: 4, oldestCursor: 1, gap: false });
  expect((await request('/logs')).body).toEqual(first.entries);
  expect((await request(`/logs?after=${first.nextCursor}`)).body.entries).toEqual([]);
  await request('/log', { message: 'next', seq: -7 });
  expect((await request('/logs?after=4')).body.entries.map(entry => entry.seq)).toEqual([5]);
});

test('retention, clear and restart preserve cursor high-water mark and disclose gaps', async () => {
  fs.writeFileSync(path.join(root, 'logs/e2e-test-logs.json'), JSON.stringify(Array.from({ length: 1000 }, (_, i) => ({ message: `legacy-${i}`, seq: 9000 }))));
  await request('/log', { message: 'new' });
  expect((await request('/logs?after=0')).body).toMatchObject({ oldestCursor: 2, nextCursor: 1001, gap: true });
  expect((await request('/logs?after=1')).body.gap).toBe(false);
  await request('/logs/clear', {});
  expect((await request('/logs?after=1000')).body).toEqual({ entries: [], oldestCursor: 1002, nextCursor: 1001, gap: true });
  expect((await request('/logs?after=1001')).body.gap).toBe(false);
  url = await start();
  expect((await request('/log', { message: 'after restart' })).body.seq).toBe(1002);
  expect((await request('/logs?after=1001')).body.entries.map(entry => entry.seq)).toEqual([1002]);
  expect((await request('/logs?after=999999')).status).toBe(409);
});

test('failed append and clear do not advance or destroy the committed cursor', async () => {
  await request('/log', { message: 'before' });
  const spy = jest.spyOn(fs, 'renameSync').mockImplementation(() => { throw new Error('failure'); });
  expect((await request('/log', { message: 'failed' })).status).toBe(500);
  expect((await request('/logs/clear', {})).status).toBe(500);
  expect((await request('/logs?after=0')).body).toMatchObject({ nextCursor: 1, gap: false });
  spy.mockRestore();
  expect((await request('/log', { message: 'after' })).body.seq).toBe(2);
});

test.each(['-1', '1.5', 'NaN', '', '9007199254740992', '0&after=1', '0&since=2020-01-01'])('rejects invalid/ambiguous cursor %s', async after => {
  expect((await request(`/logs?after=${after}`)).status).toBe(400);
});

test('CLI executes explicit status/trigger/claim/complete/pause and cursor reads over HTTP', async () => {
  expect((await cli(['status'])).body.state).toBe('paused');
  const trigger = await cli(['trigger']);
  expect(trigger.code).toBe(0);
  const runId = trigger.body.runId;
  expect((await cli(['claim', '--run-id', runId])).body.state).toBe('claimed');
  expect((await cli(['claim', '--run-id', runId])).code).toBe(1);
  expect((await cli(['complete', '--run-id', runId, '--outcome', 'passed'])).body.state).toBe('completed');
  expect((await cli(['pause'])).body.state).toBe('paused');
  const second = (await cli(['trigger'])).body.runId;
  expect((await cli(['pause', '--run-id', second])).body.state).toBe('paused');
  await request('/log', { message: 'no timestamp' });
  const logs = await cli(['get-logs', '--after', '0']);
  expect(logs.code).toBe(0);
  expect(logs.body.entries).toHaveLength(1);
  expect(logs.body.nextCursor).toBe(1);
  expect(Array.isArray((await cli(['get-logs'])).body)).toBe(true);
  expect((await cli(['status'], '')).code).toBe(1);
  expect((await cli(['status'], 'x'.repeat(40))).code).toBe(1);
});

test('CLI redacts even an echoed token and does not follow redirects', async () => {
  await request('/log', { message: TOKEN });
  const output = await cli(['get-logs']);
  expect(output.stdout).not.toContain(TOKEN);
  expect(output.stdout).toContain('[redacted]');
  const app = express();
  let requests = 0;
  app.use((_req, res) => { requests++; res.redirect(`${url}/api/e2e-loop/trigger`); });
  url = await start(app);
  const redirected = await cli(['status']);
  expect(redirected.code).toBe(1);
  expect(redirected.stderr).toContain('302');
  expect(requests).toBe(1);
});

test('byte retention reports the exact oldest sequence and leaves the complete snapshot within 4 MiB', async () => {
  const entries = Array.from({ length: 100 }, (_, i) => ({ seq: i + 1, message: 'x'.repeat(41000) }));
  fs.writeFileSync(path.join(root, 'logs/e2e-test-logs.json'), JSON.stringify({ entries, lastSequence: 100 }));
  for (let i = 0; i < 4; i++) await request('/log', { message: 'x'.repeat(60000) });
  const result = (await request('/logs?after=0')).body;
  expect(result.gap).toBe(true);
  expect(result.oldestCursor).toBe(result.entries[0].seq);
  expect(result.nextCursor).toBe(104);
  expect(fs.statSync(path.join(root, 'logs/e2e-test-logs.json')).size).toBeLessThanOrEqual(4 * 1024 * 1024);
});

test('CLI invalid arguments, malformed responses and connection errors exit nonzero', async () => {
  const missingRun = await cli(['claim']);
  expect(missingRun.code).toBe(1);
  expect(missingRun.stderr).toContain('--run-id');
  const app = express();
  app.use((_req, res) => res.send('not JSON'));
  url = await start(app);
  expect((await cli(['status'])).code).toBe(1);
  const server = servers.pop();
  await new Promise(resolve => server.close(resolve));
  expect((await cli(['status'])).code).toBe(1);
});

test('CLI refuses unsafe destinations and ambiguous commands without making requests', () => {
  expect(baseUrl().origin).toBe('https://localhost:3000');
  for (const value of ['http://example.com', 'https://example.com', 'http://127.0.0.1.evil.test', 'http://user:secret@localhost', 'http://localhost/path', 'file:///etc/passwd', 'http://localhost/?token=secret']) expect(() => baseUrl(value)).toThrow();
  expect(baseUrl('http://127.0.0.1:4000').hostname).toBe('127.0.0.1');
  expect(baseUrl('https://[::1]:4000').hostname).toBe('[::1]');
  for (const args of [[], ['claim'], ['complete', '--run-id', 'x'], ['status', '--token', TOKEN], ['trigger', '--run-id', 'x'], ['get-logs', '--after', '-1'], ['status', '--base-url', url, '--base-url', url]]) expect(() => command(args)).toThrow();
});
