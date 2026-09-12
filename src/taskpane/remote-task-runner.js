import { createProposalRecord, createProposalRuntime } from './proposal-runtime.js';

const FORBIDDEN_KEYS = /^(?:callback|on[A-Z]|range|full[-_]?text|document[-_]?text|plain[-_]?text|base64|bytes|image[-_]?bytes|ooxml|html)$/i;
const MAX_INPUT_DEPTH = 8;
const MAX_INPUT_ITEMS = 200;

function clone(value) {
    return value === undefined ? value : JSON.parse(JSON.stringify(value));
}

function isPlain(value) {
    if (!value || typeof value !== 'object') return false;
    const prototype = Object.getPrototypeOf(value);
    return prototype === Object.prototype || prototype === null;
}

function validateProposalInput(value, depth = 0, seen = new Set()) {
    if (depth > MAX_INPUT_DEPTH) throw new TypeError('Prepared proposal input is too deeply nested');
    if (typeof value === 'function' || typeof value === 'symbol' || typeof value === 'bigint') throw new TypeError('Prepared proposal input contains non-serializable data');
    if (value === null || typeof value !== 'object') {
        if (typeof value === 'string' && (/^(?:data:|base64,)/i.test(value) || value.length > 4000)) throw new TypeError('Prepared proposal input contains invalid text');
        return;
    }
    if (seen.has(value)) throw new TypeError('Prepared proposal input contains a cyclic value');
    if (!isPlain(value) && !Array.isArray(value)) throw new TypeError('Prepared proposal input contains a non-plain object');
    if (value.constructor?.name === 'Range' || ('context' in value && typeof value.load === 'function')) throw new TypeError('Word Range is not allowed in a prepared proposal');
    if (Array.isArray(value) && value.length > MAX_INPUT_ITEMS) throw new TypeError('Prepared proposal input contains too many items');
    seen.add(value);
    for (const [key, child] of Object.entries(value)) {
        if (FORBIDDEN_KEYS.test(key)) throw new TypeError(`Prepared proposal input contains forbidden field: ${key}`);
        validateProposalInput(child, depth + 1, seen);
    }
    seen.delete(value);
}

function requireIdentity(identity, name) {
    if (!isPlain(identity) || ['workspaceId', 'documentId', 'instanceId'].some((key) => typeof identity[key] !== 'string' || !identity[key])) {
        throw new TypeError(`${name} must include workspaceId, documentId, and instanceId`);
    }
    return { workspaceId: identity.workspaceId, documentId: identity.documentId, instanceId: identity.instanceId };
}

function sameParty(a, b) {
    return !!a && !!b && ['workspaceId', 'documentId', 'instanceId'].every((key) => a[key] === b[key]);
}

function matchesTarget(event, identity) {
    return sameParty(event?.target, identity);
}

function taskIdFor(event) {
    const taskId = event?.payload?.taskId;
    if (typeof taskId !== 'string' || !taskId) throw new TypeError('task.submit requires taskId');
    return taskId;
}

function proposalEventPayload(record) {
    return {
        proposalId: record.proposalId,
        taskId: record.taskId,
        revision: record.revision,
        baseRevision: String(record.baseRevision),
        artifactIds: [],
        ...(record.kind ? { kind: record.kind } : {}),
        ...(record.scope ? { scope: record.scope } : {}),
        ...(record.title ? { title: record.title } : {}),
        ...(record.summary ? { summary: record.summary } : {}),
    };
}

function proposalTerminalPayload(record) {
    return { proposalId: record.proposalId, revision: record.revision };
}

export function createRemoteTaskRunner({
    client, identity, documentAgent, prepareTask = documentAgent?.prepareTask?.bind(documentAgent), createProposal = (prepared) => prepared,
    onProposal = () => {}, onStatus = () => {}, clock = Date.now, isOnline = () => true,
} = {}) {
    if (!client || typeof client.publishEnvelope !== 'function') throw new TypeError('Remote task runner requires client.publishEnvelope');
    const localIdentity = requireIdentity(identity || client.identity, 'identity');
    if (!documentAgent || typeof documentAgent.getDocumentRevision !== 'function' || typeof documentAgent.applyProposal !== 'function') {
        throw new TypeError('Remote task runner requires documentAgent.getDocumentRevision and documentAgent.applyProposal');
    }
    if (typeof prepareTask !== 'function') throw new TypeError('Remote task runner requires prepareTask');
    const tasks = new Map();
    const proposals = new Map();
    let disposed = false;
    const status = (value) => { try { onStatus(value); } catch { /* Observer errors cannot change write outcomes. */ } };
    const failure = (code, message) => Object.assign(new Error(message), { code });
    const codeFor = (error) => /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(error?.code) ? error.code : 'REMOTE_TASK_FAILED';

    function guard(task) {
        if (disposed || !isOnline()) throw failure('offline', 'Target coordination is offline');
        if (task.controller.signal.aborted) throw failure('cancelled', 'Task was cancelled');
        if (task.event.createdAt + task.event.ttlMs <= clock()) throw failure('expired', 'Task expired');
    }
    async function publish(task, type, payload, suffix, target = task.event.source) {
        const event = task.event;
        const input = { type, target, payload, correlationId: event.correlationId, createdAt: event.createdAt, ttlMs: event.ttlMs, idempotencyKey: `${task.event.payload.taskId.slice(0, 70)}:${suffix}` };
        const response = await client.publishEnvelope(input);
        const accepted = response?.event;
        if (!accepted || accepted.type !== type || !accepted.result || !sameParty(accepted.source, localIdentity)) throw failure('invalid-response', 'Server did not acknowledge the operation');
        return accepted.result;
    }
    async function reportFailure(task, error, record) {
        const code = codeFor(error);
        try {
            if (record) await publish(task, 'proposal.conflict', { ...proposalTerminalPayload(record), reason: code }, 'conflict');
            await publish(task, 'task.failed', { taskId: task.event.payload.taskId, errorCode: code }, 'failed');
        } catch (publishError) { status({ type: 'task.failed.publish-error', taskId: task.event.payload.taskId, error: publishError }); }
    }
    function makeRuntime(record, task) {
        let lease;
        let operation = null;
        let finished = false;
        const base = createProposalRuntime(record, (ids, snapshot, options) => documentAgent.applyProposal(ids, snapshot, options), {
            identity: localIdentity, clock, finalizeSelection: true,
            isOnline: () => !disposed && isOnline(), isCancelled: () => task.controller.signal.aborted,
            getRevision: () => documentAgent.getDocumentRevision(record.scope),
            beforeApply: async () => {
                guard(task);
                lease = await publish(task, 'lease.acquire', { resourceId: 'write', durationMs: 30000 }, 'lease', localIdentity);
                checkLease();
                await publish(task, 'proposal.decision', { ...proposalTerminalPayload(record), decision: 'accepted', resourceId: 'write', fence: lease.fence }, 'accept');
            },
            beforeWrite: async () => {
                guard(task);
                lease = await publish(task, 'lease.renew', { resourceId: 'write', fence: lease.fence, durationMs: 30000 }, 'renew', localIdentity);
                const getSnapshot = client.transport?.getSnapshot || client.getSnapshot;
                if (typeof getSnapshot !== 'function') throw failure('snapshot-unavailable', 'Authoritative task state is unavailable');
                const snapshot = await getSnapshot.call(client.transport || client);
                const remoteTask = snapshot.tasks?.find((item) => item.taskId === record.taskId);
                const remoteProposal = snapshot.proposals?.find((item) => item.proposalId === record.proposalId);
                if (!remoteTask || !['claimed', 'running'].includes(remoteTask.state) || remoteTask.correlationId !== task.event.correlationId || !sameParty(remoteTask.target, localIdentity)) throw failure('task-inactive', 'Task is no longer active');
                if (remoteProposal?.state !== 'accepted' || remoteProposal.fence !== lease.fence) throw failure('proposal-inactive', 'Proposal is no longer accepted');
                checkLease();
            },
            validateWrite: () => { guard(task); checkLease(); },
        });
        function checkLease() {
            if (!lease || lease.resourceId !== 'write' || !Number.isSafeInteger(lease.fence) || lease.fence <= 0 || lease.expiresAt <= clock() || !sameParty(lease.source, localIdentity)) throw failure('lease-invalid', 'Write lease is missing or expired');
        }
        async function release() {
            if (!lease) return;
            try { await publish(task, 'lease.release', { resourceId: 'write', fence: lease.fence }, 'release', localIdentity); }
            catch (error) { status({ type: 'lease.release-error', proposalId: record.proposalId, error }); }
        }
        function review(decision, selectedIds, options = {}) {
            if (operation) return operation;
            if (finished) return Promise.resolve({ ok: false, record: base.getRecord(), conflict: { code: 'terminal', message: 'Review is already complete; retry is blocked' }, localResult: base.getLocalResult() });
            finished = true;
            operation = (async () => {
                let result;
                try {
                    guard(task);
                    if (options.signal?.aborted) throw failure('cancelled', 'Review was cancelled');
                    if (decision === 'reject') {
                        await publish(task, 'proposal.decision', { ...proposalTerminalPayload(record), decision: 'rejected' }, 'reject');
                        result = base.reject();
                    } else {
                        result = await base.apply(selectedIds, { signal: options.signal || task.controller.signal });
                        if (!result.ok) throw failure(result.conflict.code, result.conflict.message);
                        const revision = await documentAgent.getDocumentRevision(record.scope);
                        await publish(task, 'proposal.applied', { ...proposalTerminalPayload(record), documentRevision: String(revision), resourceId: 'write', fence: lease.fence }, 'applied');
                    }
                    await publish(task, 'task.succeeded', { taskId: record.taskId, proposalId: record.proposalId }, 'succeeded');
                    return { ...result, record: base.getRecord(), localResult: base.getLocalResult() };
                } catch (error) {
                    const wrote = base.getRecord().state === 'applied' || base.getRecord().state === 'unknown';
                    base.markConflict(wrote ? 'unknown' : 'conflict');
                    await reportFailure(task, error, record);
                    status({ type: wrote ? 'proposal.unknown' : 'proposal.conflict', proposalId: record.proposalId, error });
                    return { ok: false, record: base.getRecord(), localResult: base.getLocalResult(), conflict: { code: wrote ? 'write-outcome-unknown' : codeFor(error), message: String(error.message || error) } };
                } finally { await release(); }
            })();
            return operation.finally(() => { operation = null; });
        }
        return Object.freeze({ getRecord: base.getRecord, getLocalResult: base.getLocalResult, apply: (ids, options) => review('apply', ids, options), reject: () => review('reject') });
    }
    async function handleSubmit(event) {
        const taskId = taskIdFor(event);
        const key = `${event.source.instanceId}:${taskId}`;
        if (tasks.has(key)) return tasks.get(key).work;
        const task = { event: clone(event), controller: new AbortController(), work: null };
        tasks.set(key, task);
        task.work = (async () => {
            let claimed = false;
            try {
                guard(task);
                validateProposalInput(event.payload);
                await publish(task, 'task.claim', { taskId }, 'claim');
                claimed = true;
                guard(task);
                const payload = { ...event.payload, ...(event.payload.payload || {}), taskId };
                delete payload.payload;
                const prepared = await prepareTask(payload, { signal: task.controller.signal });
                validateProposalInput(prepared);
                const input = await createProposal(prepared, payload);
                validateProposalInput(input);
                guard(task);
                const expiresAt = Math.min(Number(input.expiresAt ?? Infinity), event.createdAt + event.ttlMs);
                const record = createProposalRecord({ ...input, proposalId: input.proposalId || event.payload.proposalId || `${taskId.slice(0, 110)}-proposal`, graphId: event.payload.graphId || input.graphId || 'remote-task-graph', taskId, attemptId: event.payload.attemptId || input.attemptId || `${taskId.slice(0, 110)}-attempt`, source: event.source, target: localIdentity, reviewRequired: true, state: 'pending', revision: 0, createdAt: clock(), expiresAt });
                await publish(task, 'proposal.announced', proposalEventPayload(record), 'announced');
                guard(task);
                const runtime = makeRuntime(record, task);
                proposals.set(record.proposalId, { runtime, taskId });
                try { onProposal(clone(record), runtime); } catch (error) { status({ type: 'proposal.observer-error', error }); }
                status({ type: 'proposal.announced', taskId, proposalId: record.proposalId });
            } catch (error) {
                if (claimed) await reportFailure(task, error);
                status({ type: 'task.failed', taskId, error });
            }
            return true;
        })();
        return task.work;
    }
    async function handleEvent(event) {
        if (disposed || !matchesTarget(event, localIdentity)) return false;
        if (event.type === 'task.cancel') {
            const task = tasks.get(`${event.source.instanceId}:${event.payload?.taskId}`);
            if (!task || task.event.correlationId !== event.correlationId || !sameParty(task.event.source, event.source)) return false;
            task.controller.abort();
            return true;
        }
        if (event.type !== 'task.submit') return false;
        requireIdentity(event.source, 'source');
        if (!Number.isSafeInteger(event.createdAt) || !Number.isSafeInteger(event.ttlMs) || !event.correlationId) return false;
        return handleSubmit(event);
    }
    function getProposal(proposalId) {
        const entry = proposals.get(proposalId);
        return entry ? { ...entry, record: entry.runtime.getRecord() } : null;
    }
    function dispose() {
        disposed = true;
        for (const task of tasks.values()) task.controller.abort();
    }
    return Object.freeze({ handleEvent, getProposal, dispose });
}

export { validateProposalInput };
