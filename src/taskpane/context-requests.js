const DEFAULT_TIMEOUT_MS = 30000;
const FIELDS = ['workspaceId', 'documentId', 'instanceId'];
const validId = (value) => typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(value);
const same = (a, b) => !!a && !!b && FIELDS.every((field) => a[field] === b[field]);
const validParty = (value) => !!value && FIELDS.every((field) => validId(value[field]));
const clone = (value) => JSON.parse(JSON.stringify(value));
const keyFor = (id, correlation, peer) => JSON.stringify([id, correlation, ...FIELDS.map((field) => peer[field])]);
let sequence = 0;
const newId = (prefix) => `${prefix}-${Date.now().toString(36)}-${(++sequence).toString(36)}`;

export function createContextRequestManager({ client, agent, identity = client?.identity, clock = Date.now, setTimeoutImpl = setTimeout, clearTimeoutImpl = clearTimeout, timeoutMs = DEFAULT_TIMEOUT_MS, cacheLimit = 1000, onError = () => {} } = {}) {
    if (!client || typeof client.requestContext !== 'function' || typeof client.respondContext !== 'function') throw new Error('Context request manager requires a v2 coordination client');
    if (!agent || typeof agent.readContext !== 'function' || !validParty(identity)) throw new Error('Context request manager requires a document agent and complete identity');
    if (!Number.isInteger(cacheLimit) || cacheLimit < 1 || !Number.isInteger(timeoutMs) || timeoutMs < 1) throw new Error('Invalid context limits');
    const local = Object.freeze(Object.fromEntries(FIELDS.map((field) => [field, identity[field]])));
    const pending = new Map();
    const inbound = new Map();
    const completed = new Map();
    const inboundCompleted = new Map();
    let disposed = false;

    function validEvent(event) {
        return !!event && validParty(event.source) && same(event.target, local) && event.source.workspaceId === local.workspaceId
            && event.workspaceId === local.workspaceId && validId(event.correlationId) && validId(event.payload?.requestId)
            && Number.isSafeInteger(event.createdAt) && event.createdAt >= 0 && event.createdAt <= clock() + 30000
            && Number.isInteger(event.ttlMs) && event.ttlMs >= 1000 && event.ttlMs <= 86400000 && event.createdAt + event.ttlMs > clock();
    }
    function prune() {
        const now = clock();
        for (const [key, item] of completed) if (item.expiresAt <= now) completed.delete(key);
        for (const [key, item] of inboundCompleted) if (item.expiresAt <= now) inboundCompleted.delete(key);
        for (const [key, item] of inbound) if (item.expiresAt <= now) { item.controller.abort(); clearTimeoutImpl(item.timer); inbound.delete(key); }
        while (completed.size > cacheLimit) completed.delete(completed.keys().next().value);
        while (inboundCompleted.size > cacheLimit) inboundCompleted.delete(inboundCompleted.keys().next().value);
    }
    function notifyCancel(item, reason) {
        const cancel = client.cancelContext || client.cancelRequest;
        if (typeof cancel !== 'function') return;
        void Promise.resolve().then(() => cancel.call(client, item.target, { requestId: item.requestId, reason }, {
            correlationId: item.correlationId, ttlMs: 1000, createdAt: clock(), idempotencyKey: item.cancelKey,
        })).catch(onError);
    }
    function finish(item, error, result) {
        if (!pending.delete(item.key)) return;
        clearTimeoutImpl(item.timer);
        item.signal?.removeEventListener('abort', item.abort);
        completed.set(item.key, { expiresAt: item.expiresAt });
        while (completed.size > cacheLimit) completed.delete(completed.keys().next().value);
        if (error) item.reject(error); else item.resolve(result);
    }
    function cancelItem(item, reason, code) {
        if (!pending.has(item.key)) return;
        notifyCancel(item, reason);
        finish(item, Object.assign(new Error(reason), { code, requestId: item.requestId, correlationId: item.correlationId }));
    }
    function requestContext(target, payload = {}, options = {}) {
        prune();
        if (disposed || !validParty(target) || target.workspaceId !== local.workspaceId) return Promise.reject(new Error('Invalid context target'));
        if (options.signal?.aborted) return Promise.reject(Object.assign(new Error('Context request aborted'), { name: 'AbortError' }));
        const requestId = payload.requestId ?? newId('request');
        const correlationId = options.correlationId ?? newId('correlation');
        const ttlMs = options.ttlMs ?? Math.max(1000, timeoutMs);
        if (!validId(requestId) || !validId(correlationId) || !validId(payload.scope) || !Number.isInteger(ttlMs) || ttlMs < 1000 || ttlMs > 86400000) return Promise.reject(new Error('Invalid context request'));
        const key = keyFor(requestId, correlationId, target);
        if (pending.has(key) || completed.has(key)) return Promise.reject(new Error('Context request already exists'));
        if (pending.size >= cacheLimit) return Promise.reject(new Error('Context request capacity reached'));
        const createdAt = clock();
        const body = { requestId, scope: payload.scope, ...(payload.artifactIds ? { artifactIds: clone(payload.artifactIds) } : {}) };
        const item = { key, requestId, correlationId, target: Object.freeze(clone(target)), createdAt, expiresAt: createdAt + ttlMs, signal: options.signal, cancelKey: newId('cancel') };
        const promise = new Promise((resolve, reject) => { item.resolve = resolve; item.reject = reject; });
        item.abort = () => cancelItem(item, 'Context request aborted', 'ABORTED');
        item.timer = setTimeoutImpl(() => cancelItem(item, 'Context request timed out', 'TIMEOUT'), Math.min(timeoutMs, ttlMs));
        pending.set(key, item);
        item.signal?.addEventListener('abort', item.abort, { once: true });
        try { void client.requestContext(item.target, body, { correlationId, ttlMs, createdAt, idempotencyKey: newId('request') }).catch((error) => finish(item, error)); }
        catch (error) { finish(item, error); }
        return promise;
    }
    async function sendResponse(item) {
        if (disposed || item.cancelled || item.controller.signal.aborted || item.expiresAt <= clock()) return;
        await client.respondContext(item.event, item.result, item.metadata);
    }
    async function handleEvent(event) {
        prune();
        if (disposed || !validEvent(event)) return false;
        const key = keyFor(event.payload.requestId, event.correlationId, event.source);
        if (event.type === 'context.response') {
            const item = pending.get(key);
            if (!item) return completed.has(key);
            if (item.expiresAt <= clock()) return false;
            const payload = event.payload;
            const hasSnapshot = payload.snapshot && typeof payload.snapshot === 'object' && !Array.isArray(payload.snapshot);
            const hasError = payload.error && typeof payload.error === 'object' && !Array.isArray(payload.error);
            if (!!hasSnapshot === !!hasError || Object.keys(payload).some((field) => !['requestId', 'snapshot', 'error'].includes(field))) {
                finish(item, new Error('Invalid context response'));
            } else if (hasError) {
                finish(item, Object.assign(new Error(String(payload.error.message || 'Context read failed')), { code: payload.error.code || 'CONTEXT_READ_FAILED' }));
            } else if ((payload.snapshot.sourceDocumentId !== undefined && payload.snapshot.sourceDocumentId !== item.target.documentId)
                || (payload.snapshot.sourceInstanceId !== undefined && payload.snapshot.sourceInstanceId !== item.target.instanceId)) {
                finish(item, new Error('Context snapshot identity mismatch'));
            } else finish(item, null, clone(payload));
            return true;
        }
        if (event.type === 'context.cancel') {
            const item = inbound.get(key);
            if (!item) return false;
            item.cancelled = true; item.controller.abort(); clearTimeoutImpl(item.timer);
            return true;
        }
        if (event.type !== 'context.request' || !validId(event.payload.scope)) return false;
        const cached = inboundCompleted.get(key);
        if (cached) {
            if (JSON.stringify(event.payload) !== JSON.stringify(cached.event.payload)) return false;
            await sendResponse(cached).catch(onError);
            return true;
        }
        const existing = inbound.get(key);
        if (existing) {
            if (JSON.stringify(event.payload) !== JSON.stringify(existing.event.payload) || event.createdAt !== existing.event.createdAt || event.ttlMs !== existing.event.ttlMs) return false;
            await existing.promise;
            await sendResponse(existing).catch(onError);
            return true;
        }
        if (inbound.size >= cacheLimit) { onError(new Error('Inbound context capacity reached')); return false; }
        const item = { event: clone(event), controller: new AbortController(), expiresAt: event.createdAt + event.ttlMs, cancelled: false,
            metadata: { correlationId: event.correlationId, createdAt: event.createdAt, ttlMs: event.ttlMs, idempotencyKey: newId('response') } };
        item.timer = setTimeoutImpl(() => { item.controller.abort(); inbound.delete(key); }, item.expiresAt - clock());
        inbound.set(key, item);
        item.promise = Promise.resolve().then(async () => {
            try { item.result = { requestId: event.payload.requestId, snapshot: await agent.readContext(clone(event.payload), { signal: item.controller.signal }) }; }
            catch (error) { item.result = { requestId: event.payload.requestId, error: { code: String(error.code || 'CONTEXT_READ_FAILED'), message: String(error.message || error).slice(0, 2048) } }; }
        });
        await item.promise;
        inbound.delete(key);
        clearTimeoutImpl(item.timer);
        if (!disposed) inboundCompleted.set(key, item);
        prune();
        await sendResponse(item).catch(onError);
        return true;
    }
    function dispose() {
        if (disposed) return;
        disposed = true;
        for (const item of [...pending.values()]) cancelItem(item, 'Context request manager stopped', 'ABORTED');
        for (const item of inbound.values()) { item.controller.abort(); clearTimeoutImpl(item.timer); }
        inbound.clear(); completed.clear(); inboundCompleted.clear();
    }
    return Object.freeze({ requestContext, handleEvent, dispose, pendingCount: () => pending.size });
}

export { DEFAULT_TIMEOUT_MS };
