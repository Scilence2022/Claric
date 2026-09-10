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
        ? globalThis.CLARIC_COORDINATION_TOKEN : '', onSnapshot = () => {}, onError = () => {}, mode = 'polling',
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
        await postEvent({ version: PROTOCOL_VERSION, workspaceId: identity.workspaceId, documentId: identity.documentId, type: 'presence', clientId: identity.instanceId, payload: { active: true, ephemeral: identity.documentEphemeral } }).catch(onError);
        await poll(identity);
        if (!stopped && typeof setIntervalImpl === 'function') timer = setIntervalImpl(() => { void poll(identity); }, intervalMs);
    }
    async function stop(identity) {
        stopped = true;
        stopPolling();
        if (identity) await postEvent({ version: PROTOCOL_VERSION, workspaceId: identity.workspaceId, documentId: identity.documentId, type: 'presence', clientId: identity.instanceId, payload: { active: false, ephemeral: identity.documentEphemeral } }).catch(onError);
    }
    return { getSnapshot, postEvent, poll, start, stop, stopPolling, get timer() { return timer; } };
}

export function createCoordinationClient(options = {}) {
    const identity = options.identity || createCoordinationIdentity(options);
    const transport = options.transport || createCoordinationTransport(options);
    let lastLogKey = '';
    const onSnapshot = options.onSnapshot || (() => {});
    const onError = options.onError || (() => {});
    return {
        identity,
        transport,
        async start() {
            await transport.start(identity).catch(onError);
            return identity;
        },
        async refresh() {
            const snapshot = await transport.poll(identity);
            if (snapshot) {
                const key = JSON.stringify([snapshot.sequence, Object.keys(snapshot.presence).length, snapshot.metadata]);
                if (key !== lastLogKey) { lastLogKey = key; onSnapshot(snapshot); }
            } else onError(new Error('Coordination snapshot unavailable'));
            return snapshot;
        },
        stop: () => transport.stop(identity),
        record: (type, payload) => transport.postEvent({ version: PROTOCOL_VERSION, workspaceId: identity.workspaceId, documentId: identity.documentId, type, clientId: identity.instanceId, payload }),
        _onSnapshot: onSnapshot,
        _onError: onError,
    };
}

export { DEFAULT_POLL_INTERVAL_MS, PROTOCOL_VERSION };
