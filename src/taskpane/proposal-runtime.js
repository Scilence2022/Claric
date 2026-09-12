const RECORD_FIELDS = ['proposalId', 'graphId', 'taskId', 'attemptId', 'source', 'target', 'kind', 'scope', 'title', 'summary', 'items', 'baseRevision', 'revision', 'state', 'reviewRequired', 'createdAt', 'expiresAt'];
const TERMINAL_STATES = new Set(['applied', 'rejected', 'conflict', 'unknown']);
const MAX_TEXT = 4000;
const MAX_ITEMS = 200;
const MAX_ITEM_KEYS = new Set(['id', 'label', 'before', 'after', 'status']);

function plain(value) {
    if (!value || typeof value !== 'object') return false;
    const proto = Object.getPrototypeOf(value);
    return proto === Object.prototype || proto === null;
}

function assertSafe(value, seen = new Set()) {
    if (typeof value === 'function' || typeof value === 'symbol' || typeof value === 'bigint') throw new TypeError('Proposal record contains non-serializable data');
    if (value === null || typeof value !== 'object') return;
    if (seen.has(value)) throw new TypeError('Proposal record contains a cyclic value');
    if (!plain(value) && !Array.isArray(value)) throw new TypeError('Proposal record contains a non-plain object');
    if (value.constructor?.name === 'Range' || ('context' in value && typeof value.load === 'function')) throw new TypeError('Word Range is not allowed in a proposal record');
    seen.add(value);
    for (const child of Array.isArray(value) ? value : Object.values(value)) assertSafe(child, seen);
    seen.delete(value);
}

function text(value, name, optional = true) {
    if (value === undefined && optional) return undefined;
    if (typeof value !== 'string' || value.length > MAX_TEXT || /^(data:|base64,)/i.test(value)) throw new TypeError(`Invalid or oversized ${name}`);
    return value;
}

function identity(value, name) {
    if (!plain(value)) throw new TypeError(`${name} must be a plain object`);
    const result = {};
    for (const key of ['workspaceId', 'documentId', 'instanceId']) {
        if (value[key] !== undefined) {
            if (typeof value[key] !== 'string' || value[key].length > 256) throw new TypeError(`Invalid ${name}.${key}`);
            result[key] = value[key];
        }
    }
    if (!result.documentId || !result.instanceId) throw new TypeError(`${name} must include documentId and instanceId`);
    return result;
}

function normalizeItems(items) {
    if (!Array.isArray(items) || items.length > MAX_ITEMS) throw new TypeError('Proposal items must be a bounded array');
    return items.map((item, index) => {
        if (!plain(item) || item.id === undefined || (typeof item.id !== 'string' && typeof item.id !== 'number')) throw new TypeError(`Invalid proposal item ${index}`);
        const result = { id: String(item.id) };
        for (const key of MAX_ITEM_KEYS) if (key !== 'id' && item[key] !== undefined) {
            if (key === 'status') {
                if (!['pending', 'applied', 'skipped'].includes(item[key])) throw new TypeError('Invalid proposal item status');
                result[key] = item[key];
            } else result[key] = text(item[key], `item.${key}`);
        }
        return result;
    });
}

function normalize(input) {
    if (!plain(input)) throw new TypeError('Proposal input must be a plain object');
    assertSafe(input);
    const result = {};
    for (const key of RECORD_FIELDS) {
        if (input[key] !== undefined) result[key] = input[key];
    }
    for (const key of ['proposalId', 'graphId', 'taskId', 'attemptId', 'kind', 'scope']) result[key] = text(input[key], key, false);
    result.source = identity(input.source, 'source');
    result.target = identity(input.target, 'target');
    result.title = text(input.title, 'title', false);
    result.summary = text(input.summary, 'summary', false);
    result.items = normalizeItems(input.items || []);
    if (input.baseRevision === undefined || (typeof input.baseRevision !== 'string' && typeof input.baseRevision !== 'number')) throw new TypeError('baseRevision is required');
    result.baseRevision = input.baseRevision;
    if (input.revision !== undefined && (!Number.isSafeInteger(input.revision) || input.revision < 0)) throw new TypeError('Invalid proposal revision');
    result.revision = input.revision ?? 0;
    result.state = input.state || 'pending';
    if (!['pending', 'applying', 'applied', 'rejected', 'conflict', 'unknown'].includes(result.state)) throw new TypeError('Invalid proposal state');
    result.reviewRequired = input.reviewRequired !== false;
    for (const key of ['createdAt', 'expiresAt']) {
        if (typeof input[key] !== 'string' && typeof input[key] !== 'number') throw new TypeError(`${key} is required`);
        result[key] = input[key];
    }
    return result;
}

export function serializeProposalRecord(record) {
    const normalized = normalize(record);
    return JSON.parse(JSON.stringify(normalized));
}

export function createProposalRecord(input) {
    return serializeProposalRecord({ ...input, state: input.state || 'pending' });
}

export function reduceProposalRecord(record, decision, selectedItemIds) {
    const next = serializeProposalRecord(record);
    if (TERMINAL_STATES.has(next.state)) return next;
    if (decision === 'reject') return { ...next, state: 'rejected' };
    if (decision !== 'apply') throw new TypeError('Unknown proposal decision');
    const selected = selectedItemIds == null ? next.items.map((item) => item.id) : selectedItemIds.map(String);
    next.items = next.items.map((item) => selected.includes(item.id) ? { ...item, status: 'applied' } : item);
    const applied = next.items.filter((item) => item.status === 'applied').length;
    next.state = applied === next.items.length ? 'applied' : 'applying';
    return next;
}

function conflict(code, message, details = {}) { return { ok: false, conflict: { code, message, ...details } }; }

export function createProposalRuntime(record, applyImpl, options = {}) {
    if (typeof applyImpl !== 'function') throw new TypeError('applyImpl must be a function');
    let current = serializeProposalRecord(record);
    let inFlight = null;
    let localResult;
    const clock = options.clock || Date.now;
    const identityValue = options.identity || options.document || options.target;
    const getRevision = options.getRevision;
    function getRecord() { return serializeProposalRecord(current); }
    function updateRecord(next) {
        if (inFlight || TERMINAL_STATES.has(current.state)) throw new Error('Cannot replace an active or terminal proposal');
        current = serializeProposalRecord(next);
        return getRecord();
    }
    function getLocalResult() { return localResult; }
    function markConflict(state = 'conflict') {
        if (!['conflict', 'unknown'].includes(state)) throw new TypeError('Invalid conflict state');
        current = { ...current, state };
        return getRecord();
    }
    function ownershipConflict() {
        if (!identityValue) return null;
        for (const key of ['workspaceId', 'documentId', 'instanceId']) if (identityValue[key] !== undefined && current.target[key] !== identityValue[key]) return conflict('target-mismatch', `Proposal target ${key} does not belong to this runtime`, { field: key });
        return null;
    }
    function terminalConflict(code, message, details = {}, state = 'conflict') {
        current = { ...current, state };
        return conflict(code, message, { ...details, state });
    }
    async function apply(selectedItemIds, applyOptions = {}) {
        if (inFlight) return inFlight;
        inFlight = (async () => {
            if (TERMINAL_STATES.has(current.state)) return conflict('terminal', 'Proposal has already reached a terminal state', { state: current.state });
            const owned = ownershipConflict(); if (owned) return owned;
            if (!current.reviewRequired) return conflict('review-required', 'Proposal is not eligible for interactive review');
            if (options.online === false || options.isOnline?.() === false) return conflict('offline', 'Remote coordination or document is offline');
            if (options.isCancelled?.() === true) return conflict('cancelled', 'The remote task was cancelled');
            if (Number(current.expiresAt) <= clock()) return conflict('expired', 'Proposal has expired');
            if (applyOptions.signal?.aborted) return conflict('cancelled', 'Proposal apply was cancelled');
            if (typeof getRevision !== 'function') return conflict('revision-unavailable', 'Current document revision is unavailable');
            const revision = await getRevision();
            if (String(revision) !== String(current.baseRevision)) return conflict('stale', 'Document revision no longer matches proposal', { expected: current.baseRevision, actual: revision });
            const ids = selectedItemIds == null ? undefined : selectedItemIds.map(String);
            const pending = current.items.filter((item) => item.status !== 'applied').map((item) => item.id);
            const selected = ids === undefined ? pending : ids.filter((id) => pending.includes(id));
            if (!selected.length) return conflict('no-items', 'No pending proposal items were selected');
            if (typeof options.beforeApply === 'function') {
                const before = await options.beforeApply(getRecord(), { signal: applyOptions.signal });
                if (before === false) return conflict('cancelled', 'Proposal apply was cancelled');
                if (before && before.ok === false) return before;
            }
            if (typeof options.beforeWrite === 'function') await options.beforeWrite(getRecord());
            const checkedRevision = await getRevision();
            if (String(checkedRevision) !== String(current.baseRevision)) return conflict('stale', 'Document revision changed before write', { expected: current.baseRevision, actual: checkedRevision });
            if (options.isCancelled?.() || applyOptions.signal?.aborted) return conflict('cancelled', 'Proposal apply was cancelled');
            if (options.isOnline?.() === false) return conflict('offline', 'Remote coordination is offline');
            if (Number(current.expiresAt) <= clock()) return conflict('expired', 'Proposal has expired');
            if (typeof options.validateWrite === 'function') options.validateWrite(getRecord());
            current = { ...current, state: 'applying' };
            try {
                localResult = await applyImpl([...new Set(selected)], getRecord(), { signal: applyOptions.signal });
                if (localResult?.ok === false) throw Object.assign(new Error('Document agent did not confirm the write'), { localResult });
            } catch (error) {
                if (error.localResult !== undefined) localResult = error.localResult;
                return terminalConflict('apply-unknown', 'The document write outcome is unknown; retry is blocked', { errorCode: String(error?.code || 'APPLY_UNKNOWN') }, 'unknown');
            }
            current = reduceProposalRecord(current, 'apply', selected);
            if (options.finalizeSelection) current = { ...current, state: 'applied', items: current.items.map((item) => item.status === 'applied' ? item : { ...item, status: 'skipped' }) };
            return { ok: true, record: getRecord(), appliedItemIds: selected, localResult };
        })();
        try { return await inFlight; }
        catch (error) { return terminalConflict(String(error.code || 'preflight-failed'), String(error.message || error)); }
        finally { inFlight = null; }
    }
    function reject() {
        if (inFlight) return conflict('busy', 'An apply operation is in progress');
        if (TERMINAL_STATES.has(current.state)) return { ok: false, conflict: { code: 'terminal', message: 'Proposal has already reached a terminal state', state: current.state } };
        const owned = ownershipConflict(); if (owned) return owned;
        current = reduceProposalRecord(current, 'reject');
        return { ok: true, record: getRecord() };
    }
    return Object.freeze({ getRecord, updateRecord, getLocalResult, markConflict, apply, reject });
}
