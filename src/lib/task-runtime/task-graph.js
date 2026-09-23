import { normalizeCompound, TASK_STATES } from './task-model.js';
import { createResourceLocks } from './resource-locks.js';
import { TASK_EVENTS, createEventSink } from './task-events.js';

export function validateTaskGraph(graph) {
    const normalized = normalizeCompound(graph);
    const ids = new Set();
    const errors = [];
    normalized.tasks.forEach((task) => { if (ids.has(task.taskId)) errors.push(`duplicate taskId: ${task.taskId}`); ids.add(task.taskId); });
    normalized.tasks.forEach((task) => task.dependsOn.forEach((dep) => { if (!ids.has(dep)) errors.push(`${task.taskId} depends on missing task ${dep}`); }));
    const visiting = new Set(); const visited = new Set();
    function visit(id) { if (visiting.has(id)) return true; if (visited.has(id)) return false; visiting.add(id); const task = normalized.tasks.find((item) => item.taskId === id); const cycle = task && task.dependsOn.some(visit); visiting.delete(id); visited.add(id); return cycle; }
    normalized.tasks.forEach((task) => { if (visit(task.taskId)) errors.push('task graph contains a dependency cycle'); });
    return { valid: errors.length === 0, errors, graph: normalized };
}

/**
 * @typedef {object} TaskGraphExecuteOptions
 * @property {(event: { type: string, [key: string]: any }) => void} [onEvent]
 *   Callback invoked for each task-lifecycle event (GRAPH_STARTED,
 *   TASK_STARTED, TASK_SUCCEEDED / TASK_FAILED / GRAPH_FINISHED).
 * @property {AbortSignal} [signal]
 *   Optional signal; when aborted, in-flight tasks are cancelled and the
 *   graph stops dispatching new work.
 * @property {Map<string, any>} [initialResults]
 *   Successful results from an earlier pass. Blocked tasks are deliberately
 *   omitted so they can be reconsidered after a proposal is applied.
 */

/**
 * Runs the dependency-respecting dispatch loop: every pass picks all
 * `READY` tasks whose dependencies have `SUCCEEDED`, runs them under a
 * resource lock, and emits lifecycle events. Detects cycles and missing
 * dependencies up front via {@link validateTaskGraph}.
 *
 * @param {object} input - Compound graph payload ({ graphId?, tasks })
 * @param {(task: object, ctx: { signal?: AbortSignal, inputs: any[] }) => Promise<any>} execute
 * @param {TaskGraphExecuteOptions} [options]
 * @returns {Promise<{ graph: object, results: Map<string, any> }>}
 */
export async function executeTaskGraph(input, execute, { onEvent = undefined, signal = undefined, initialResults = undefined } = {}) {
    const checked = validateTaskGraph(input);
    if (!checked.valid) throw new Error(checked.errors.join('; '));
    const graph = checked.graph;
    const locks = createResourceLocks();
    const emit = createEventSink(onEvent);
    const results = new Map();
    const artifacts = new Map();
    if (initialResults) {
        if (!(initialResults instanceof Map)) throw new Error('initialResults must be a Map');
        const taskIds = new Set(graph.tasks.map((task) => task.taskId));
        for (const [id, result] of initialResults) {
            if (!taskIds.has(id) || result?.state !== TASK_STATES.SUCCEEDED) {
                throw new Error(`Cannot resume an invalid task result: ${id}`);
            }
            results.set(id, result);
            if (Array.isArray(result.value?.artifacts)) {
                result.value.artifacts.forEach((artifact, index) => artifacts.set(`${id}:${index}`, artifact));
            }
        }
    }
    const checkAbort = () => { if (signal?.aborted) throw new DOMException('Task graph cancelled', 'AbortError'); };
    const successful = (result) => result?.state === TASK_STATES.SUCCEEDED
        && (result.value?.status !== 'no_op' || result.value.satisfied === true);
    function blocked(task, message) {
        task.state = TASK_STATES.BLOCKED;
        results.set(task.taskId, { state: task.state, error: new Error(message) });
        emit(TASK_EVENTS.TASK_BLOCKED, { graphId: graph.graphId, taskId: task.taskId, error: message });
    }
    emit(TASK_EVENTS.GRAPH_STARTED, { graphId: graph.graphId });
    while (results.size < graph.tasks.length) {
        checkAbort();
        for (const task of graph.tasks) {
            if (!results.has(task.taskId) && task.dependsOn.some((id) => results.has(id) && !successful(results.get(id)))) {
                blocked(task, 'A required task failed, was blocked, or did not satisfy its request.');
            }
        }
        const ready = graph.tasks.filter((task) => !results.has(task.taskId)
            && task.dependsOn.every((id) => successful(results.get(id))) && locks.canAcquire(task.resources, task.taskId));
        if (!ready.length) {
            if (results.size === graph.tasks.length) break;
            // A blocked dependency may have appeared after its successor in the array.
            if (graph.tasks.some((t) => !results.has(t.taskId) && t.dependsOn.some((id) => results.has(id) && !successful(results.get(id))))) continue;
            throw new Error('Task graph is blocked');
        }
        for (const task of ready) {
            checkAbort();
            locks.acquire(task.resources, task.taskId);
            task.state = TASK_STATES.RUNNING;
            const identity = { graphId: graph.graphId, taskId: task.taskId, attemptId: task.attemptId };
            emit(TASK_EVENTS.TASK_STARTED, identity);
            try {
                const inputs = task.dependsOn.map((id) => ({ taskId: id, ...results.get(id) }));
                if (task.inputRefs.some((id) => !artifacts.has(id))) {
                    blocked(task, 'A required task artifact is unavailable.');
                    continue;
                }
                inputs.push(...task.inputRefs.map((id) => ({ artifactId: id, value: artifacts.get(id) })));
                const value = await execute(task, { signal, inputs });
                checkAbort();
                if (value?.status === 'failed') throw (value.error instanceof Error ? value.error : new Error(value.error || 'Task failed'));
                if (value?.status === 'blocked') { blocked(task, value.error?.message || 'Task is blocked'); continue; }
                if (value?.status === 'no_op' && value.satisfied !== true) {
                    throw new Error('Task produced no change without confirming that the request is already satisfied.');
                }
                task.state = TASK_STATES.SUCCEEDED;
                results.set(task.taskId, { state: task.state, value });
                if (Array.isArray(value?.artifacts)) value.artifacts.forEach((artifact, index) => artifacts.set(`${task.taskId}:${index}`, artifact));
                const event = value?.status === 'staged' ? TASK_EVENTS.TASK_STAGED
                    : value?.status === 'no_op' ? TASK_EVENTS.TASK_NO_OP : TASK_EVENTS.TASK_SUCCEEDED;
                emit(event, { ...identity, value });
            } catch (error) {
                if (signal?.aborted || error?.name === 'AbortError') {
                    task.state = TASK_STATES.CANCELLED;
                    results.set(task.taskId, { state: task.state, error });
                    throw error;
                }
                task.state = TASK_STATES.FAILED;
                results.set(task.taskId, { state: task.state, error });
                emit(TASK_EVENTS.TASK_FAILED, { ...identity, error: error.message });
            } finally { locks.release(task.resources, task.taskId); }
        }
    }
    emit(TASK_EVENTS.GRAPH_FINISHED, { graphId: graph.graphId });
    return { graph, results };
}
