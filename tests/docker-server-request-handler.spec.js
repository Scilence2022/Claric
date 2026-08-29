/**
 * Regression tests for the request boundary of docker-server.cjs.
 *
 * The crash-class regression covered here: a decoded NUL byte in the
 * request path (GET /%00) made fs.readFile throw synchronously inside the
 * request handler — one unauthenticated request took the whole server
 * process (and every in-flight LLM stream) down. The request boundary must
 * (1) reject control characters up front, (2) answer unexpected throws with
 * a per-connection 500, and (3) keep the process alive so a follow-up
 * request still succeeds.
 */

const http = require('http');
const { Readable, Writable } = require('stream');
const { requestHandler, configureServer } = require('../scripts/docker-server.cjs');

/**
 * Minimal IncomingMessage stand-in: handleRequest reads method/url/headers
 * and (on the proxy path) pipes the body.
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
 * Minimal ServerResponse stand-in capturing writeHead/setHeader/end; the
 * `done` promise resolves when the response stream finishes.
 */
function makeRes() {
    const chunks = [];
    const res = new Writable({
        write(chunk, _enc, cb) { chunks.push(chunk); cb(); },
    });
    res.headersSent = false;
    res.statusCode = undefined;
    res.headers = {};
    res.setHeader = (name, value) => { res.headers[name.toLowerCase()] = value; };
    res.writeHead = (statusCode, headers) => {
        res.statusCode = statusCode;
        // Real ServerResponse merges headers set via setHeader with the
        // writeHead overrides (writeHead wins on conflicts) and lowercases
        // names; mirror that so tests read headers the way production
        // responses expose them.
        const merged = { ...res.headers };
        for (const [name, value] of Object.entries(headers || {})) {
            merged[name.toLowerCase()] = value;
        }
        res.headers = merged;
        res.headersSent = true;
        return res;
    };
    const done = new Promise((resolve) => res.on('finish', () => resolve({
        statusCode: res.statusCode,
        body: Buffer.concat(chunks).toString(),
    })));
    return { res, done };
}

describe('docker-server request boundary validation', () => {
    test('rejects a decoded NUL byte in the path with 400, without throwing', async () => {
        // Before the fix, routing this path reached fs.readFile, whose
        // synchronous ERR_INVALID_ARG_VALUE killed the process.
        const { res, done } = makeRes();
        expect(() => requestHandler(makeReq('GET', '/%00'), res)).not.toThrow();
        const result = await done;
        expect(result.statusCode).toBe(400);
    });

    test('rejects other decoded control characters (newline, tab, DEL)', async () => {
        for (const raw of ['/%0A', '/%09', '/%7F']) {
            const { res, done } = makeRes();
            requestHandler(makeReq('GET', raw), res);
            const result = await done;
            expect(result.statusCode).toBe(400);
        }
    });

    test('still rejects path traversal', async () => {
        const { res, done } = makeRes();
        requestHandler(makeReq('GET', '/../package.json'), res);
        const result = await done;
        expect(result.statusCode).toBe(400);
    });

    test('missing static assets stay a plain 404', async () => {
        const { res, done } = makeRes();
        requestHandler(makeReq('GET', '/no-such-asset-xyz.png'), res);
        const result = await done;
        expect(result.statusCode).toBe(404);
    });

    test('non-GET/HEAD on a non-proxy path is 405 with an Allow header', async () => {
        const { res, done } = makeRes();
        requestHandler(makeReq('POST', '/whatever'), res);
        const result = await done;
        expect(result.statusCode).toBe(405);
        expect(res.headers.allow).toBe('GET, HEAD');
    });
});

describe('docker-server process resilience', () => {
    let server;
    let port;

    beforeAll(async () => {
        // A real loopback server wired exactly like production (minus TLS),
        // so the regression is exercised through the actual async dispatch
        // path instead of a direct call.
        server = http.createServer(requestHandler);
        await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
        port = server.address().port;
    });

    afterAll(async () => {
        await new Promise((resolve) => server.close(resolve));
    });

    test('a NUL-path request does not kill the server', async () => {
        const crash = await fetch(`http://127.0.0.1:${port}/%00`);
        expect(crash.status).toBe(400);

        // The process survived: the very next request is answered normally.
        const alive = await fetch(`http://127.0.0.1:${port}/healthz`);
        expect(alive.status).toBe(200);
        expect((await alive.json()).status).toBe('ok');

        const again = await fetch(`http://127.0.0.1:${port}/%00`);
        expect(again.status).toBe(400);
    });

    test('configureServer registers process-level crash guards', () => {
        const dummy = configureServer(http.createServer(() => {}));
        expect(process.listenerCount('uncaughtException')).toBeGreaterThanOrEqual(1);
        expect(process.listenerCount('unhandledRejection')).toBeGreaterThanOrEqual(1);
        dummy.close();
    });
});
