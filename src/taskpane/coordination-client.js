const PROTOCOL_VERSION = 1;
const DEFAULT_POLL_INTERVAL_MS = 10000;

function randomId(prefix) {
    const random = typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function'
        ? crypto.randomUUID()
        : `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    return `${prefix}-${random}`.slice(0, 128);
}

function stableId(prefix, value) {
    const text = String(value || '').trim();
    if (!text) return { id: randomId(prefix), ephemeral: true };
    let hash = 2166136261;
    for (let index = 0; index < text.length; index += 1) hash = Math.imul(hash ^ text.charCodeAt(index), 16777619);
    return { id: `${prefix}-${(hash >>> 0).toString(16)}`, ephemeral: false };
}

export function createCoordinationIdentity({ office = typeof Office !== 'undefined' ? Office : null, locationObject = typeof location !== 'undefined' ? location : null, storage = typeof localStorage !== 'undefined' ? localStorage : null } = {}) {
    const origin = locationObject && locationObject.origin ? locationObject.origin : '';
    let workspaceSource = origin;
    try {
        if (storage) {
            const key = 'claric.coordination.workspace';
            workspaceSource = storage.getItem(key) || origin || randomId('workspace-source');
            storage.setItem(key, workspaceSource);
        }
    } catch {
        workspaceSource = origin;
    }
    const workspace = stableId('workspace', workspaceSource);
    const documentSource = office && office.context && office.context.document && (
        office.context.document.url || office.context.document.id || office.context.document.key
    );
    const document = stableId('document', documentSource);
    return {
        instanceId: randomId('instance'),
        workspaceId: workspace.id,
        documentId: document.id,
        workspaceEphemeral: workspace.ephemeral,
        documentEphemeral: document.ephemeral,
    };
}

export function filterCoordinationSnapshot(snapshot, identity) {
    if (!snapshot || snapshot.workspaceId !== identity.workspaceId) return null;
    const presence = {};
    for (const [clientId, value] of Object.entries(snapshot.presence || {})) presence[clientId] = { ...value };
    const metadata = {
        ...(snapshot.metadataByDocument && snapshot.metadataByDocument['*'] || {}),
        ...(snapshot.metadataByDocument && snapshot.metadataByDocument[identity.documentId] || {}),
    };
    const events = (snapshot.events || [])
        .filter((event) => !event.documentId || event.documentId === identity.documentId)
        .map((event) => ({ ...event, payload: { ...(event.payload || {}) } }));
    return { ...snapshot, presence, metadata, events };
}

export function createCoordinationTransport({
    baseUrl = typeof globalThis !== 'undefined' && globalThis.CLARIC_COORDINATION_URL
        ? globalThis.CLARIC_COORDINATION_URL : '/coordination', fetchImpl = typeof fetch !== 'undefined' ? fetch.bind(globalThis) : null,
    intervalMs = DEFAULT_POLL_INTERVAL_MS, setIntervalImpl = typeof setInterval !== 'undefined' ? setInterval : null,
    clearIntervalImpl = typeof clearInterval !== 'undefined' ? clearInterval : null, token = typeof globalThis !== 'undefined' && globalThis.CLARIC_COORDINATION_TOKEN
        ? globalThis.CLARIC_COORDINATION_TOKEN : '', onSnapshot = () => {}, onError = () => {}, mode = 'polling', readOnly = false,
} = {}) {
    if (mode !== 'polling' && mode !== 'websocket') throw new Error(`Unsupported coordination transport: ${mode}`);
    let timer = null;
    let stopped = true;
    const headers = () => ({ ...(token ? { Authorization: `Bearer ${token}` } : {}) });
    async function request(path, options = {}) {
        if (typeof fetchImpl !== 'function') throw new Error('Coordination fetch is unavailable');
        const response = await fetchImpl(`${baseUrl}${path}`, { ...options, headers: { ...headers(), ...(options.headers || {}) } });
        let body = null;
        try { body = await response.json(); } catch { /* server may return an empty error */ }
        if (!response.ok) throw new Error(body && body.error ? body.error : `Coordination request failed (${response.status})`);
        return body;
    }
    async function getSnapshot(identity) {
        const query = encodeURIComponent(identity.workspaceId);
        return request(`/snapshot?workspaceId=${query}`);
    }
    function postEvent(event) {
        if (readOnly) return Promise.reject(new Error('Legacy coordination fallback is read-only'));
        return request('/events', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(event) });
    }
    function stopPolling() {
        if (timer !== null && typeof clearIntervalImpl === 'function') clearIntervalImpl(timer);
        timer = null;
    }
    async function poll(identity) {
        try {
            const snapshot = filterCoordinationSnapshot(await getSnapshot(identity), identity);
            if (snapshot) onSnapshot(snapshot);
            return snapshot;
        } catch (error) {
            onError(error);
            return null;
        }
    }
    async function start(identity) {
        stopPolling();
        stopped = false;
        if (!readOnly) await postEvent({ version: PROTOCOL_VERSION, workspaceId: identity.workspaceId, documentId: identity.documentId, type: 'presence', clientId: identity.instanceId, payload: { active: true, ephemeral: identity.documentEphemeral } });
        const snapshot = await poll(identity);
        if (!snapshot) throw new Error('Coordination snapshot unavailable');
        if (!stopped && typeof setIntervalImpl === 'function') timer = setIntervalImpl(() => { void poll(identity); }, intervalMs);
    }
    async function stop(identity) {
        stopped = true;
        stopPolling();
        if (identity && !readOnly) await postEvent({ version: PROTOCOL_VERSION, workspaceId: identity.workspaceId, documentId: identity.documentId, type: 'presence', clientId: identity.instanceId, payload: { active: false, ephemeral: identity.documentEphemeral } }).catch(onError);
    }
    return { getSnapshot, postEvent, poll, start, stop, stopPolling, get timer() { return timer; } };
}

const PROTOCOL_V2 = 2;
const DEFAULT_ENVELOPE_TTL_MS = 60000;
const MIN_ENVELOPE_TTL_MS = 1000;
const MAX_ENVELOPE_TTL_MS = 86400000;
const PARTY_FIELDS = ['workspaceId', 'documentId', 'instanceId'];

function party(value) {
    if (!value || PARTY_FIELDS.some((field) => typeof value[field] !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(value[field]))) {
        throw new Error('Invalid coordination identity');
    }
    return Object.freeze(Object.fromEntries(PARTY_FIELDS.map((field) => [field, value[field]])));
}

function sameParty(a, b) {
    return !!a && !!b && PARTY_FIELDS.every((field) => a[field] === b[field]);
}

function validateEnvelopeMetadata(input, now) {
    for (const field of ['correlationId', 'idempotencyKey']) {
        if (typeof input[field] !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(input[field])) throw new Error(`Invalid ${field}`);
    }
    if (!Number.isSafeInteger(input.ttlMs) || input.ttlMs < MIN_ENVELOPE_TTL_MS || input.ttlMs > MAX_ENVELOPE_TTL_MS) throw new Error('Invalid TTL');
    if (!Number.isSafeInteger(input.createdAt) || input.createdAt < 0 || input.createdAt > now + 30000 || input.createdAt + input.ttlMs <= now) throw new Error('Expired or invalid createdAt');
}

export function createV2CoordinationTransport({
    baseUrl = globalThis.CLARIC_COORDINATION_URL || '/coordination',
    fetchImpl = typeof fetch !== 'undefined' ? fetch.bind(globalThis) : null,
    token = globalThis.CLARIC_COORDINATION_TOKEN || '', intervalMs = DEFAULT_POLL_INTERVAL_MS,
    setIntervalImpl = setInterval, clearIntervalImpl = clearInterval,
    requestTimeoutMs = 15000, setTimeoutImpl = setTimeout, clearTimeoutImpl = clearTimeout,
    onSnapshot = () => {}, onEvents = () => {}, onError = () => {}, mode = 'polling', clock = Date.now,
} = {}) {
    if (mode !== 'polling' && mode !== 'websocket') throw new Error(`Unsupported coordination transport: ${mode}`);
    let credential = '';
    let identity = null;
    let expiresAt = null;
    let cursor = null;
    let epoch = null;
    let snapshot = null;
    let timer = null;
    let stopped = true;
    let polling = null;
    let starting = null;
    let ticking = null;
    const requests = new Set();
    const root = baseUrl.replace(/\/$/, '');

    async function request(path, body, registering = false, signal) {
        if (typeof fetchImpl !== 'function') throw new Error('Coordination fetch is unavailable');
        if (!registering && !credential) throw new Error('Coordination instance is not registered');
        const authorization = registering ? token : credential;
        const controller = new AbortController();
        requests.add(controller);
        const abort = () => controller.abort();
        const timeout = setTimeoutImpl(abort, requestTimeoutMs);
        let abortListener;
        const aborted = new Promise((resolve, reject) => {
            abortListener = () => reject(Object.assign(new Error('Coordination request aborted or timed out'), { name: 'AbortError' }));
            controller.signal.addEventListener('abort', abortListener, { once: true });
        });
        signal?.addEventListener('abort', abort, { once: true });
        if (signal?.aborted) abort();
        try {
            return await Promise.race([aborted, (async () => {
                const response = await fetchImpl(`${root}/v2${path}`, {
                    method: body === undefined ? 'GET' : 'POST',
                    headers: { ...(authorization ? { Authorization: `Bearer ${authorization}` } : {}), ...(body === undefined ? {} : { 'Content-Type': 'application/json' }) },
                    ...(body === undefined ? {} : { body: JSON.stringify(body) }), signal: controller.signal,
                });
                if (!response.ok) throw Object.assign(new Error(`Coordination request failed (${response.status})`), { status: response.status, registering });
                return response.json();
            })()]);
        } finally {
            clearTimeoutImpl(timeout);
            signal?.removeEventListener('abort', abort);
            controller.signal.removeEventListener('abort', abortListener);
            requests.delete(controller);
        }
    }
    function updateCursor(value) {
        if (!value || !Number.isSafeInteger(value.cursor) || value.cursor < 0 || typeof value.epoch !== 'string' || !value.epoch) throw new Error('Invalid coordination cursor');
        cursor = value.cursor;
        epoch = value.epoch;
    }
    function snapshotShape(value) {
        if (!value || value.version !== PROTOCOL_V2 || value.workspaceId !== identity.workspaceId || !Array.isArray(value.documents)) throw new Error('Invalid coordination snapshot');
        const presence = Object.fromEntries(value.documents.map((entry) => [entry.instanceId, { ...entry, clientId: entry.instanceId }]));
        return { ...value, sequence: value.cursor, presence, metadata: {}, events: [] };
    }
    function recoveryEvents(value) {
        const events = [];
        for (const record of value.requests || []) {
            if (record.expiresAt <= clock()) continue;
            const response = record.state === 'responded';
            const type = response ? 'context.response' : record.state === 'cancelled' ? 'context.cancel' : 'context.request';
            const payload = response ? record.result : type === 'context.request'
                ? { requestId: record.requestId, scope: record.scope, ...(record.artifactIds ? { artifactIds: record.artifactIds } : {}) }
                : { requestId: record.requestId, reason: 'Cancelled' };
            events.push({ version: 2, workspaceId: record.workspaceId, source: response ? record.target : record.source,
                target: response ? record.source : record.target, type, payload, correlationId: record.correlationId,
                idempotencyKey: `recovery-${record.requestId}`.slice(0, 128), createdAt: record.createdAt,
                ttlMs: record.expiresAt - record.createdAt, recovered: true });
        }
        return events.filter((event) => sameParty(event.source, identity) || sameParty(event.target, identity));
    }
    async function recover(value) {
        const fresh = snapshotShape(value);
        const events = recoveryEvents(fresh);
        if (!stopped) {
            await onSnapshot(fresh);
            if (events.length) await onEvents(events);
        }
        updateCursor(value);
        snapshot = { ...fresh, events };
        return snapshot;
    }
    async function getSnapshot() {
        return snapshotShape(await request('/snapshot'));
    }
    async function refreshDiscovery(previousCursor, previousEpoch) {
        const value = await request('/snapshot');
        const fresh = snapshotShape(value);
        const epochChanged = previousEpoch !== null && fresh.epoch !== previousEpoch;
        if (epochChanged) return { snapshot: fresh, epochChanged };
        cursor = previousCursor; epoch = previousEpoch;
        snapshot = { ...fresh, cursor, epoch, sequence: cursor };
        return { snapshot, epochChanged };
    }
    async function publishEnvelope(input) {
        if (!identity || !credential || stopped) throw new Error('Coordination instance is not registered');
        if (input.source && !sameParty(input.source, identity)) throw new Error('Cannot override coordination source');
        const target = party(input.target);
        if (target.workspaceId !== identity.workspaceId) throw new Error('Coordination target workspace mismatch');
        const envelope = {
            version: PROTOCOL_V2, workspaceId: identity.workspaceId, source: identity, target,
            type: input.type, payload: input.payload || {},
            correlationId: input.correlationId ?? randomId('correlation'),
            idempotencyKey: input.idempotencyKey ?? randomId('idempotency'),
            ttlMs: input.ttlMs ?? DEFAULT_ENVELOPE_TTL_MS, createdAt: input.createdAt ?? clock(),
        };
        validateEnvelopeMetadata(envelope, clock());
        return request('/envelopes', envelope, false, input.signal);
    }
    function stopPolling() {
        if (timer !== null && typeof clearIntervalImpl === 'function') clearIntervalImpl(timer);
        timer = null;
    }
    async function pollOnce() {
        const previousCursor = cursor;
        const previousEpoch = epoch;
        try {
            const discovery = await refreshDiscovery(previousCursor, previousEpoch);
            if (discovery.epochChanged) return recover(discovery.snapshot);
            if (!stopped) await onSnapshot(discovery.snapshot);
            let result;
            try { result = await request(`/events?after=${cursor}&epoch=${encodeURIComponent(epoch)}`); }
            catch (error) {
                if (error.status !== 410) throw error;
                return recover(await getSnapshot());
            }
            if (result.epoch !== epoch) return recover(await getSnapshot());
            if (result.cursor < cursor || !Array.isArray(result.events)) throw new Error('Invalid coordination events');
            const events = result.events.filter((event) => sameParty(event.source, identity) || sameParty(event.target, identity));
            if (!stopped && events.length) await onEvents(events);
            updateCursor(result);
            snapshot = { ...snapshot, cursor, epoch, sequence: cursor, events };
            return snapshot;
        } catch (error) {
            if (!stopped) onError(error);
            return null;
        }
    }
    function poll() {
        if (stopped || cursor === null) return Promise.resolve(null);
        if (!polling) polling = pollOnce().finally(() => { polling = null; });
        return polling;
    }
    async function heartbeat() {
        if (!stopped && identity && credential) await publishEnvelope({ type: 'document.heartbeat', target: identity, payload: {} });
    }
    async function tick() {
        if (!ticking) ticking = Promise.all([heartbeat().catch(onError), poll()]).finally(() => { ticking = null; });
        return ticking;
    }
    async function startSession(localIdentity) {
        stopPolling();
        stopped = false;
        credential = ''; identity = null; expiresAt = null; cursor = null; epoch = null; snapshot = null;
        const registration = await request('/instances/register', { workspaceId: localIdentity.workspaceId, documentId: localIdentity.documentId }, true);
        if (!registration || typeof registration.credential !== 'string' || !registration.credential || !Number.isSafeInteger(registration.expiresAt) || registration.expiresAt <= clock()) throw new Error('Invalid coordination registration');
        const issuedIdentity = party(registration.identity);
        if (issuedIdentity.workspaceId !== localIdentity.workspaceId || issuedIdentity.documentId !== localIdentity.documentId) throw new Error('Coordination registration scope mismatch');
        credential = registration.credential;
        identity = issuedIdentity;
        expiresAt = registration.expiresAt;
        await publishEnvelope({ type: 'document.announce', target: identity, payload: {} });
        const fresh = await getSnapshot();
        if (!stopped) {
            await recover(fresh);
            if (typeof setIntervalImpl === 'function') timer = setIntervalImpl(() => { void tick(); }, Math.min(intervalMs, 20000));
        }
        return identity;
    }
    function start(localIdentity) {
        if (!stopped && identity && !starting) return Promise.resolve(identity);
        if (!starting) starting = startSession(localIdentity).catch((error) => { stopped = true; stopPolling(); credential = ''; identity = null; throw error; }).finally(() => { starting = null; });
        return starting;
    }
    async function stop() {
        stopPolling();
        const leaving = !stopped && credential && identity;
        stopped = true;
        for (const controller of requests) controller.abort();
        if (leaving) {
            try { await request('/envelopes', { version: 2, workspaceId: identity.workspaceId, source: identity, target: identity,
                type: 'document.left', payload: {}, correlationId: randomId('left'), idempotencyKey: randomId('left'), createdAt: clock(), ttlMs: 1000 }); }
            catch (error) { onError(error); }
        }
        credential = ''; identity = null; expiresAt = null; cursor = null; epoch = null; snapshot = null;
    }
    function postEvent(event) {
        if (event.version === PROTOCOL_V2) return publishEnvelope(event);
        if (event.type === 'presence') return publishEnvelope({ type: event.payload.active ? 'document.announce' : 'document.left', target: identity, payload: {} });
        return Promise.reject(new Error('Legacy metadata is unavailable in coordination v2'));
    }
    return {
        start, stop, dispose: stop, poll, getSnapshot, postEvent, publishEnvelope, stopPolling,
        version: PROTOCOL_V2,
        get identity() { return identity; }, get expiresAt() { return expiresAt; },
        get cursor() { return cursor; }, get epoch() { return epoch; }, get timer() { return timer; },
    };
}

export function createCoordinationClient(options = {}) {
    const localIdentity = options.identity || createCoordinationIdentity(options);
    let transport = options.transport || createV2CoordinationTransport(options);
    let lastLogKey = '';
    let stopped = false;
    const onSnapshot = options.onSnapshot || (() => {});
    const onError = options.onError || (() => {});
    const currentIdentity = () => transport.identity || localIdentity;
    const requireV2 = () => transport.version === PROTOCOL_V2
        ? null : Promise.reject(new Error('Cross-document coordination requires v2'));
    const publishEnvelope = (envelope) => transport.version === PROTOCOL_V2 && typeof transport.publishEnvelope === 'function'
        ? transport.publishEnvelope(envelope) : Promise.reject(new Error('Cross-document coordination requires v2'));
    const send = (type, target, payload, metadata = {}) => publishEnvelope({ ...metadata, type, target, payload });
    return {
        get identity() { return currentIdentity(); },
        get transport() { return transport; },
        async start() {
            stopped = false;
            try {
                await transport.start(localIdentity);
            } catch (error) {
                if (!stopped && !options.transport && error.registering && [404, 501].includes(error.status)) {
                    transport.stopPolling();
                    transport = createCoordinationTransport({ ...options, readOnly: true });
                    await transport.start(localIdentity);
                } else {
                    onError(error);
                    throw error;
                }
            }
            return currentIdentity();
        },
        async refresh() {
            const snapshot = await transport.poll(currentIdentity());
            if (snapshot && transport.version !== PROTOCOL_V2) {
                const key = JSON.stringify([snapshot.sequence, Object.keys(snapshot.presence || {}).length, snapshot.metadata]);
                if (key !== lastLogKey) { lastLogKey = key; onSnapshot(snapshot); }
            } else if (!snapshot) onError(new Error('Coordination snapshot unavailable'));
            return snapshot;
        },
        stop: () => { stopped = true; return transport.stop(currentIdentity()); },
        record: (type, payload) => {
            if (transport.version === PROTOCOL_V2) {
                const identity = currentIdentity();
                return transport.postEvent({ version: PROTOCOL_VERSION, workspaceId: identity.workspaceId, documentId: identity.documentId, type, clientId: identity.instanceId, payload });
            }
            return requireV2();
        },
        publishEnvelope,
        requestContext: (target, payload, metadata) => send('context.request', target, payload, metadata),
        submitTask: (target, payload, metadata) => send('task.submit', target, payload, metadata),
        cancelContext: (target, payload, metadata) => send('context.cancel', target, payload, metadata),
        cancelTask: (target, payload, metadata) => send('task.cancel', target, payload, metadata),
        cancelRequest: (target, payload, metadata) => send('context.cancel', target, payload, metadata),
        respondContext: (request, payload, metadata = {}) => send('context.response', request.source,
            { ...payload, requestId: request.payload.requestId },
            { correlationId: request.correlationId, ttlMs: request.ttlMs, createdAt: request.createdAt, ...metadata }),
        _onSnapshot: onSnapshot,
        _onError: onError,
    };
}

export { DEFAULT_POLL_INTERVAL_MS, PROTOCOL_VERSION, PROTOCOL_V2, DEFAULT_ENVELOPE_TTL_MS };
