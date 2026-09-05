const http = require('http');
const https = require('https');

const USAGE = 'Usage: node scripts/dev-harness-client.cjs status|trigger|pause [--run-id ID]|claim --run-id ID|complete --run-id ID --outcome passed|failed|get-logs [--after SEQUENCE] [--base-url URL]';

function baseUrl(value = 'https://localhost:3000') {
  let url;
  try { url = new URL(value); } catch (_error) { throw new Error('Invalid base URL'); }
  if (!['http:', 'https:'].includes(url.protocol) || !['localhost', '127.0.0.1', '[::1]'].includes(url.hostname) || url.username || url.password || url.pathname !== '/' || url.search || url.hash) {
    throw new Error('Base URL must be a loopback HTTP(S) origin without credentials, path, query or fragment');
  }
  return url;
}

function command(args) {
  const [name, ...rest] = args;
  if (!['status', 'trigger', 'pause', 'claim', 'complete', 'get-logs'].includes(name)) throw new Error(USAGE);
  const flags = {};
  for (let i = 0; i < rest.length; i += 2) {
    const key = rest[i];
    if (!['--base-url', '--run-id', '--outcome', '--after'].includes(key) || flags[key] !== undefined || !rest[i + 1] || rest[i + 1].startsWith('--')) throw new Error(USAGE);
    flags[key] = rest[i + 1];
  }
  const allowed = ['--base-url'];
  if (['pause', 'claim', 'complete'].includes(name)) allowed.push('--run-id');
  if (name === 'complete') allowed.push('--outcome');
  if (name === 'get-logs') allowed.push('--after');
  if (Object.keys(flags).some(flag => !allowed.includes(flag))) throw new Error(USAGE);
  if (['claim', 'complete'].includes(name) && !flags['--run-id']?.trim()) throw new Error('Explicit --run-id is required');
  if (flags['--run-id'] !== undefined && !flags['--run-id'].trim()) throw new Error('Explicit --run-id must not be empty');
  if (name === 'complete' && !['passed', 'failed'].includes(flags['--outcome'])) throw new Error('--outcome must be passed or failed');
  if (flags['--after'] !== undefined && (!/^(0|[1-9]\d*)$/.test(flags['--after']) || !Number.isSafeInteger(Number(flags['--after'])))) throw new Error('--after must be a nonnegative safe integer');
  const endpoint = name === 'get-logs' ? `/logs${flags['--after'] === undefined ? '' : `?after=${flags['--after']}`}` : `/api/e2e-loop/${name}`;
  const body = ['status', 'get-logs'].includes(name) ? undefined : {
    ...(flags['--run-id'] ? { runId: flags['--run-id'] } : {}),
    ...(flags['--outcome'] ? { outcome: flags['--outcome'] } : {})
  };
  return { url: baseUrl(flags['--base-url']), endpoint, body };
}

function send({ url, endpoint, body }, token) {
  return new Promise((resolve, reject) => {
    const data = body === undefined ? undefined : JSON.stringify(body);
    const transport = url.protocol === 'https:' ? https : http;
    // Pin localhost to loopback rather than trusting a host resolver or proxy environment.
    const hostname = url.hostname === '[::1]' ? '::1' : '127.0.0.1';
    const req = transport.request({
      protocol: url.protocol,
      hostname,
      port: url.port || (url.protocol === 'https:' ? 443 : 80),
      servername: url.hostname === 'localhost' ? 'localhost' : undefined,
      rejectUnauthorized: true,
      method: data === undefined ? 'GET' : 'POST',
      path: endpoint,
      agent: false,
      headers: { Host: url.host, 'x-claric-harness-token': token, ...(data === undefined ? {} : { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) }) }
    }, res => {
      let bytes = 0;
      const chunks = [];
      res.on('data', chunk => {
        bytes += chunk.length;
        if (bytes > 5 * 1024 * 1024) req.destroy(new Error('Response exceeds size limit'));
        else chunks.push(chunk);
      });
      res.on('error', () => reject(new Error('Response interrupted')));
      res.on('end', () => {
        if (res.statusCode < 200 || res.statusCode >= 300) return reject(new Error(`Harness HTTP ${res.statusCode}; command was not acknowledged`));
        let value;
        try { value = JSON.parse(Buffer.concat(chunks).toString('utf8')); } catch (_error) { return reject(new Error('Harness returned invalid JSON')); }
        if (!value || typeof value !== 'object' || value.success === false) return reject(new Error('Harness returned an unsuccessful response'));
        resolve(value);
      });
    });
    const deadline = setTimeout(() => req.destroy(new Error('Request timed out')), 10000);
    req.on('close', () => clearTimeout(deadline));
    req.on('error', () => reject(new Error('Harness request failed (connection, TLS, timeout or response limit); no automatic retry')));
    req.end(data);
  });
}

async function main(args = process.argv.slice(2), env = process.env, io = process) {
  const token = env.CLARIC_HARNESS_TOKEN;
  const redact = value => {
    const text = String(value);
    return token ? text.split(token).join('[redacted]').split(JSON.stringify(token).slice(1, -1)).join('[redacted]') : text;
  };
  try {
    const request = command(args);
    if (!token || !/^[\x21-\x7e]{32,256}$/.test(token)) throw new Error('CLARIC_HARNESS_TOKEN environment variable is required (32-256 non-space ASCII characters)');
    const result = await send(request, token);
    io.stdout.write(`${redact(JSON.stringify(result))}\n`);
    return 0;
  } catch (error) {
    io.stderr.write(`${redact(error.message)}\n`);
    return 1;
  }
}

if (require.main === module) main().then(code => { process.exitCode = code; });

module.exports = { main, baseUrl, command };
