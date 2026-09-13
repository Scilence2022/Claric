import { createId } from '../lib/task-runtime/task-model.js';
import { validateTaskGraph } from '../lib/task-runtime/task-graph.js';

const TERMINAL = new Set(['succeeded', 'failed', 'cancelled', 'blocked']);
const EVENTS = new Set(['task.claim', 'task.progress', 'task.succeeded', 'task.failed', 'task.cancel', 'proposal.announced', 'proposal.updated', 'proposal.decision', 'proposal.applied', 'proposal.conflict']);
const FIELDS = ['workspaceId', 'documentId', 'instanceId'];
const validId = (value) => typeof value === 'string' && /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(value);
const clone = (value) => JSON.parse(JSON.stringify(value));
const same = (a, b) => !!a && !!b && FIELDS.every((field) => a[field] === b[field]);
function freezeParty(value) {
    if (!value || FIELDS.some((field) => !validId(value[field]))) throw new Error('Task target must include workspaceId, documentId, and instanceId');
    return Object.freeze(Object.fromEntries(FIELDS.map((field) => [field, value[field]])));
}

export function createDistributedTaskRuntime({ client, identity = client?.identity, resolveTarget, onStatus = () => {}, onResult = () => {}, clock = Date.now, maxGraphs = 100, maxTasks = 1000, maxEvents = 1000, maxSeenEvents = 5000 } = {}) {
    if (!client || typeof client.submitTask !== 'function' || typeof client.cancelTask !== 'function') throw new Error('Distributed runtime requires submit and task-cancel coordination APIs');
    const local = freezeParty(identity);
    if ([maxGraphs, maxTasks, maxEvents, maxSeenEvents].some((limit) => !Number.isInteger(limit) || limit < 1)) throw new Error('Invalid runtime limits');
    const graphs = new Map();
    const seen = new Map();
    let disposed = false;

    function settleGraph(graph) {
        let changed;
        do {
            changed = false;
            for (const task of graph.tasks.values()) {
                if (task.state === 'pending' && task.dependsOn.some((id) => ['failed', 'cancelled', 'blocked'].includes(graph.tasks.get(id)?.state))) { task.state = 'blocked'; changed = true; }
            }
        } while (changed);
        const tasks = [...graph.tasks.values()];
        if (graph.state !== 'cancelled' && tasks.every((task) => TERMINAL.has(task.state))) graph.state = tasks.every((task) => task.state === 'succeeded') ? 'succeeded' : 'failed';
    }
    async function dispatchReady(graph) {
        for (const task of graph.tasks.values()) {
            if (disposed || graph.state === 'cancelled') break;
            if (task.state !== 'pending' || task.issued || !task.dependsOn.every((id) => graph.tasks.get(id)?.state === 'succeeded')) continue;
            task.issued = true;
            task.state = 'submitted';
            task.createdAt = clock();
            const payload = { graphId: graph.graphId, taskId: task.taskId, attemptId: task.attemptId, taskType: task.type, instruction: task.instruction, artifactIds: task.inputRefs, reviewRequired: true };
            task.submitting = client.submitTask(task.target, payload, { correlationId: task.correlationId, idempotencyKey: task.submitKey, createdAt: task.createdAt, ttlMs: task.ttlMs });
            try {
                await task.submitting;
                onStatus({ type: 'task.submitted', graphId: graph.graphId, taskId: task.taskId, target: clone(task.target) });
            } catch (error) {
                if (!task.cancelRequested) task.state = 'failed';
                task.error = { code: error.code || 'SUBMIT_FAILED', message: error.message };
                graph.state = 'failed';
                onResult({ type: 'task.failed', graphId: graph.graphId, taskId: task.taskId, payload: clone(task.error) });
                settleGraph(graph);
                throw error;
            } finally { task.submitting = null; }
        }
        settleGraph(graph);
    }
    async function submitGraph(input, resolver = resolveTarget) {
        if (disposed) throw new Error('Distributed runtime is disposed');
        const checked = validateTaskGraph(input);
        if (!checked.valid) throw new Error(checked.errors.join('; '));
        const graph = checked.graph;
        if (!validId(graph.graphId)) throw new Error('Invalid graphId');
        if (graphs.has(graph.graphId)) throw new Error('Graph already exists');
        if (graphs.size >= maxGraphs || graph.tasks.length > maxTasks) throw new Error('Distributed graph capacity reached');
        const state = { graphId: graph.graphId, state: 'submitted', tasks: new Map(), events: [] };
        for (const task of graph.tasks) {
            if (![task.taskId, task.attemptId, task.type].every(validId) || task.instruction.length > 2048 || task.inputRefs.length > 100 || !task.inputRefs.every(validId)) throw new Error('Invalid distributed task');
            const target = freezeParty(typeof resolver === 'function' ? resolver(task) : task.target);
            if (target.workspaceId !== local.workspaceId) throw new Error('Task target workspace mismatch');
            const ttlMs = task.ttlMs ?? 60000;
            if (!Number.isInteger(ttlMs) || ttlMs < 1000 || ttlMs > 86400000) throw new Error('Invalid task TTL');
            state.tasks.set(task.taskId, { ...clone(task), target, state: 'pending', issued: false, ttlMs, correlationId: createId('correlation'), submitKey: createId('submit'), cancelKey: createId('cancel'), proposal: null });
        }
        graphs.set(graph.graphId, state);
        await dispatchReady(state);
        onStatus({ type: 'graph.submitted', graphId: graph.graphId, taskCount: graph.tasks.length });
        return getSnapshot(graph.graphId);
    }
    function locate(event) {
        for (const graph of graphs.values()) for (const task of graph.tasks.values()) {
            if (task.correlationId !== event.correlationId) continue;
            if (event.payload.taskId && event.payload.taskId !== task.taskId) continue;
            if (event.payload.graphId && event.payload.graphId !== graph.graphId) continue;
            if (event.payload.attemptId && event.payload.attemptId !== task.attemptId) continue;
            if (!event.payload.taskId && (!event.payload.proposalId || task.proposal?.proposalId !== event.payload.proposalId)) continue;
            return { graph, task };
        }
        return null;
    }
    function handleEvent(event) {
        if (disposed || !event || !EVENTS.has(event.type) || !event.payload || event.workspaceId !== local.workspaceId) return false;
        if (!Number.isSafeInteger(event.createdAt) || !Number.isInteger(event.ttlMs) || event.ttlMs < 1000 || event.createdAt + event.ttlMs <= clock() || event.createdAt > clock() + 30000) return false;
        const found = locate(event);
        if (!found) return false;
        const { graph, task } = found;
        const cancelEcho = event.type === 'task.cancel' && same(event.source, local) && same(event.target, task.target);
        if (!cancelEcho && (!same(event.target, local) || !same(event.source, task.target))) return false;
        if (event.type === 'task.cancel' && !cancelEcho) return false;
        const key = JSON.stringify([event.source.instanceId, event.idempotencyKey || event.sequence, event.type, event.correlationId]);
        if (seen.has(key)) return true;
        if (!task.issued || TERMINAL.has(task.state) || task.cancelRequested) return false;
        if (event.type === 'task.claim' && task.state !== 'submitted') return false;
        if (event.type === 'task.succeeded' && task.state === 'submitted') return false;
        if (event.type.startsWith('proposal.')) {
            const p = event.payload;
            if (!validId(p.proposalId) || !Number.isInteger(p.revision) || p.revision < 0) return false;
            if (event.type === 'proposal.announced') {
                if (task.proposal && task.proposal.proposalId !== p.proposalId) return false;
                task.proposal = { proposalId: p.proposalId, revision: p.revision, state: 'announced' };
            } else {
                if (!task.proposal || task.proposal.proposalId !== p.proposalId) return false;
                if (event.type === 'proposal.updated') {
                    if (p.revision !== task.proposal.revision + 1) return false;
                    task.proposal = { ...task.proposal, revision: p.revision, state: 'announced' };
                } else {
                    if (p.revision !== task.proposal.revision) return false;
                    task.proposal.state = event.type === 'proposal.decision' ? p.decision : event.type.split('.')[1];
                }
            }
        }
        if (event.type === 'task.succeeded' && task.proposal && (event.payload.proposalId !== task.proposal.proposalId || !['applied', 'rejected'].includes(task.proposal.state))) return false;
        seen.set(key, event.createdAt + event.ttlMs);
        for (const [id, expiresAt] of seen) if (expiresAt <= clock()) seen.delete(id);
        while (seen.size > maxSeenEvents) seen.delete(seen.keys().next().value);
        const next = { 'task.claim': 'claimed', 'task.progress': 'running', 'task.succeeded': 'succeeded', 'task.failed': 'failed', 'task.cancel': 'cancelled' }[event.type];
        if (next) task.state = next;
        task.lastEvent = clone(event);
        graph.events.push(clone(event));
        if (graph.events.length > maxEvents) graph.events.shift();
        onStatus({ type: event.type, graphId: graph.graphId, taskId: task.taskId, state: task.state });
        if (TERMINAL.has(task.state)) onResult({ type: event.type, graphId: graph.graphId, taskId: task.taskId, state: task.state, payload: clone(event.payload) });
        settleGraph(graph);
        if (event.type === 'task.succeeded') void dispatchReady(graph).catch((error) => onStatus({ type: 'graph.dispatch-failed', graphId: graph.graphId, error }));
        return true;
    }
    async function cancelTask(graphId, taskId, reason = 'Cancelled by user') {
        const graph = graphs.get(graphId); const task = graph?.tasks.get(taskId);
        if (!task) throw new Error('Unknown distributed task');
        if (TERMINAL.has(task.state)) return getSnapshot(graphId);
        task.cancelRequested = true;
        task.state = 'cancelled';
        settleGraph(graph);
        if (task.issued) {
            if (task.submitting) await task.submitting;
            await client.cancelTask(task.target, { taskId, reason }, { correlationId: task.correlationId, idempotencyKey: task.cancelKey, createdAt: clock(), ttlMs: 1000 });
        }
        return getSnapshot(graphId);
    }
    async function cancelGraph(graphId, reason = 'Cancelled by user') {
        const graph = graphs.get(graphId); if (!graph) throw new Error('Unknown distributed graph');
        graph.state = 'cancelled';
        const active = [...graph.tasks.values()].filter((task) => !TERMINAL.has(task.state));
        await Promise.all(active.map((task) => cancelTask(graphId, task.taskId, reason)));
        return getSnapshot(graphId);
    }
    function getSnapshot(graphId) {
        const graph = graphs.get(graphId);
        return graph ? clone({ graphId: graph.graphId, state: graph.state, tasks: [...graph.tasks.values()].map((task) => { const snapshot = { ...task }; delete snapshot.submitting; return snapshot; }), events: graph.events }) : null;
    }
    function dispose() { disposed = true; graphs.clear(); seen.clear(); }
    return Object.freeze({ submitGraph, handleEvent, cancelTask, cancelGraph, getSnapshot, dispose, graphIds: () => [...graphs.keys()] });
}
