/**
 * Regression tests for the LLM proxy path validation in docker-server.cjs.
 *
 * The upstream URL is built with `new URL(suffix, route.targetUrl)`, and a
 * protocol-relative suffix ("//host" — or "/\host", which WHATWG URL parsing
 * treats identically for special schemes) rebases onto an attacker-chosen
 * authority: one crafted request turned the same-origin LLM proxy into an
 * open SSRF relay. handleProxyRequest must reject any suffix that resolves
 * outside the route's configured origin, and pass same-origin paths through
 * unchanged.
 *
 * The tests use real loopback servers: an "upstream" standing in for the
 * configured LLM backend and a "decoy" listener representing the host an
 * attacker wants to reach. A regression (or an incomplete fix) shows up as
 * a connection arriving at the decoy.
 */

const http = require('http');
const { Readable, Writable } = require('stream');
const {
    buildProxyRoutes,
    handleProxyRequest,
    parseProxyTarget,
} = require('../scripts/docker-server.cjs');

describe('docker-server proxy target validation', () => {
    test.each([
        'https://api.example.com',
        'http://localhost:11434',
        'http://127.0.0.1:8026',
        'http://[::1]:8026',
        'http://host.docker.internal:11434',
    ])('allows %s', (target) => {
        expect(parseProxyTarget(target)).toBeInstanceOf(URL);
    });

    test.each([
        'http://remote.example/v1',
        'http://localhost.evil/v1',
        'ftp://localhost/model',
        'javascript:alert(1)',
        'https://user:password@example.com/v1',
        '//remote.example/v1',
    ])('rejects %s', (target) => {
        expect(parseProxyTarget(target)).toBeNull();
    });

    test('buildProxyRoutes drops remote cleartext targets without logging the URL', () => {
        const consoleError = jest.spyOn(console, 'error').mockImplementation(() => {});
        const leakedTarget = 'http://remote.example/v1?token=secret';
        try {
            const routes = buildProxyRoutes({
                OLLAMA_PROXY_PATH: '/ollama',
                OLLAMA_PROXY_TARGET: 'http://localhost:11434',
                VLLM_PROXY_PATH: '/vllm',
                VLLM_PROXY_TARGET: leakedTarget,
                OPENAI_PROXY_PATH: '/openai',
                OPENAI_PROXY_TARGET: 'https://api.openai.com',
                LLM_PROXY_TIMEOUT_MS: 300000,
            });

            expect(routes.map((route) => route.proxyPath)).toEqual(['/ollama', '/openai']);
            expect(routes[0].targetUrl.href).toBe('http://localhost:11434/');
            const log = consoleError.mock.calls.flat().join(' ');
            expect(log).toContain('VLLM_PROXY_TARGET');
            expect(log).not.toContain(leakedTarget);
            expect(log).not.toContain('secret');
        } finally {
            consoleError.mockRestore();
        }
    });
});

function startRecordingServer() {
    const hits = [];
    const server = http.createServer((req, res) => {
        const chunks = [];
        req.on('data', (c) => chunks.push(c));
        req.on('end', () => {
            hits.push({ url: req.url, host: req.headers.host, body: Buffer.concat(chunks).toString() });
            res.writeHead(200, { 'Content-Type': 'text/plain' });
            res.end('upstream-ok');
        });
    });
    return new Promise((resolve) => {
        server.listen(0, '127.0.0.1', () => resolve({ server, hits, port: server.address().port }));
    });
}

/**
 * Minimal IncomingMessage stand-in: handleProxyRequest only reads
 * method/url/headers and pipes the body into the upstream request.
 */
function makeReq(method, url, headers = {}) {
    const req = new Readable({ read() {} });
    req.method = method;
    req.url = url;
    req.headers = headers;
    req.push(null);
    return req;
}

/**
 * Minimal ServerResponse stand-in capturing writeHead/end; `done` resolves
 * when the response stream finishes (both sendError and proxied responses).
 */
function makeRes() {
    const chunks = [];
    const res = new Writable({
        write(chunk, _enc, cb) { chunks.push(chunk); cb(); },
    });
    res.headersSent = false;
    res.statusCode = undefined;
    res.headers = undefined;
    res.writeHead = (statusCode, headers) => {
        res.statusCode = statusCode;
        // Real ServerResponse lowercases header names; mirror that so tests
        // can read headers the way the production response handles expose them.
        res.headers = Object.fromEntries(
            Object.entries(headers || {}).map(([name, value]) => [name.toLowerCase(), value])
        );
        res.headersSent = true;
        return res;
    };
    const done = new Promise((resolve) => res.on('finish', () => resolve({
        statusCode: res.statusCode,
        body: Buffer.concat(chunks).toString(),
    })));
    return { res, done };
}

describe('docker-server LLM proxy path validation', () => {
    let upstream;
    let decoy;
    const routes = [];

    beforeAll(async () => {
        upstream = await startRecordingServer();
        decoy = await startRecordingServer();
    });

    afterAll(async () => {
        for (const route of routes) {
            if (route.agent) route.agent.destroy();
        }
        await Promise.all([
            new Promise((resolve) => upstream.server.close(resolve)),
            new Promise((resolve) => decoy.server.close(resolve)),
        ]);
    });

    function makeRoute() {
        const route = {
            proxyPath: '/ollama',
            targetUrl: new URL(`http://127.0.0.1:${upstream.port}`),
            timeoutMs: 2000,
            agent: null,
        };
        routes.push(route);
        return route;
    }

    test('rejects a protocol-relative suffix that would rebase to another host', async () => {
        const maliciousPath = `/ollama//127.0.0.1:${decoy.port}/steal`;
        const { res, done } = makeRes();

        handleProxyRequest(makeRoute(), makeReq('GET', maliciousPath), res, maliciousPath);

        // Clear the guard timer once settled so no handle is left open for jest.
        let guard;
        const result = await Promise.race([
            done,
            new Promise((_, reject) => { guard = setTimeout(() => reject(new Error('response never settled')), 2000); }),
        ]).finally(() => clearTimeout(guard));
        expect(result.statusCode).toBe(400);
        expect(result.body).toBe('Invalid proxy path');

        // Give a hypothetical buggy outbound connection a moment to land.
        await new Promise((resolve) => setTimeout(resolve, 100));
        expect(decoy.hits).toEqual([]);
        expect(upstream.hits).toEqual([]);
    });

    test('rejects the backslash variant ("/\\host") of the same rebasing', async () => {
        const maliciousPath = `/ollama/\\127.0.0.1:${decoy.port}/steal`;
        const { res, done } = makeRes();

        handleProxyRequest(makeRoute(), makeReq('GET', maliciousPath), res, maliciousPath);

        const result = await done;
        expect(result.statusCode).toBe(400);

        await new Promise((resolve) => setTimeout(resolve, 100));
        expect(decoy.hits).toEqual([]);
    });

    test('still proxies a normal same-origin path to the configured upstream', async () => {
        const path = '/ollama/v1/models';
        const { res, done } = makeRes();

        handleProxyRequest(makeRoute(), makeReq('GET', path, { authorization: 'Bearer sk-test' }), res, path);

        const result = await done;
        expect(result.statusCode).toBe(200);
        expect(result.body).toBe('upstream-ok');

        expect(upstream.hits).toHaveLength(1);
        expect(upstream.hits[0].url).toBe('/v1/models');
        // The Host header is rewritten to the configured target, not the
        // request path's authority.
        expect(upstream.hits[0].host).toBe(`127.0.0.1:${upstream.port}`);
    });

    test('rejects oversized declared bodies with 413', async () => {
        const route = makeRoute();
        const { res, done } = makeRes();
        const headers = { 'content-length': String(33 * 1024 * 1024) };
        const hitsBefore = upstream.hits.length;

        handleProxyRequest(route, makeReq('POST', '/ollama/v1/chat/completions', headers), res, '/ollama/v1/chat/completions');

        const result = await done;
        expect(result.statusCode).toBe(413);
        expect(result.body).toBe('Request body too large');
        // The oversized request never reached the upstream.
        expect(upstream.hits).toHaveLength(hitsBefore);
    });

    test('reports a generic upstream failure without leaking the cause', async () => {
        // Nothing listens on port 1: the connect attempt fails fast with
        // ECONNREFUSED, which must stay in the server log — not in the
        // response body that reveals local topology.
        const route = {
            proxyPath: '/ollama',
            targetUrl: new URL('http://127.0.0.1:1'),
            timeoutMs: 2000,
            agent: null,
        };
        routes.push(route);
        const { res, done } = makeRes();

        handleProxyRequest(route, makeReq('POST', '/ollama/v1/chat/completions'), res, '/ollama/v1/chat/completions');

        const result = await done;
        expect(result.statusCode).toBe(502);
        expect(result.body).toBe('LLM upstream unavailable');
        expect(result.body).not.toContain('ECONNREFUSED');
    });

    test('answers the CORS preflight without touching any upstream', async () => {
        const { res, done } = makeRes();

        handleProxyRequest(makeRoute(), makeReq('OPTIONS', '/ollama/v1/chat/completions'), res, '/ollama/v1/chat/completions');

        const result = await done;
        expect(result.statusCode).toBe(204);
        expect(res.headers['access-control-allow-methods']).toContain('POST');

        await new Promise((resolve) => setTimeout(resolve, 50));
        expect(upstream.hits).toHaveLength(1); // only the pass-through test above
    });
});
