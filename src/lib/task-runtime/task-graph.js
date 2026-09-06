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

export async function executeTaskGraph(input, execute, { onEvent, signal } = {}) {
    const checked = validateTaskGraph(input); if (!checked.valid) throw new Error(checked.errors.join('; '));
    const graph = checked.graph; const locks = createResourceLocks(); const emit = createEventSink(onEvent); const results = new Map();
    emit(TASK_EVENTS.GRAPH_STARTED, { graphId: graph.graphId });
    while (results.size < graph.tasks.length) {
        if (signal?.aborted) throw new Error('Task graph cancelled');
        const ready = graph.tasks.filter((task) => !results.has(task.taskId) && task.dependsOn.every((id) => results.get(id)?.state === TASK_STATES.SUCCEEDED) && locks.canAcquire(task.resources, task.taskId));
        if (!ready.length) { const unresolved = graph.tasks.filter((task) => !results.has(task.taskId)); if (unresolved.length) throw new Error('Task graph is blocked'); break; }
        for (const task of ready) { locks.acquire(task.resources, task.taskId); task.state = TASK_STATES.RUNNING; emit(TASK_EVENTS.TASK_STARTED, { graphId: graph.graphId, taskId: task.taskId, attemptId: task.attemptId }); try { const value = await execute(task, { signal }); task.state = TASK_STATES.SUCCEEDED; results.set(task.taskId, { state: task.state, value }); emit(TASK_EVENTS.TASK_SUCCEEDED, { graphId: graph.graphId, taskId: task.taskId, attemptId: task.attemptId, value }); } catch (error) { if (signal?.aborted || error?.name === 'AbortError') { task.state = TASK_STATES.CANCELLED; results.set(task.taskId, { state: task.state, error }); throw error; } task.state = TASK_STATES.FAILED; results.set(task.taskId, { state: task.state, error }); emit(TASK_EVENTS.TASK_FAILED, { graphId: graph.graphId, taskId: task.taskId, attemptId: task.attemptId, error: error.message }); } finally { locks.release(task.resources, task.taskId); } }
    }
    emit(TASK_EVENTS.GRAPH_FINISHED, { graphId: graph.graphId }); return { graph, results };
}
