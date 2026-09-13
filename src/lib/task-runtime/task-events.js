/** Structured task lifecycle events. */
export const TASK_EVENTS = Object.freeze({ GRAPH_STARTED: 'task.graph.started', TASK_READY: 'task.ready', TASK_STARTED: 'task.started', TASK_SUCCEEDED: 'task.succeeded', TASK_FAILED: 'task.failed', GRAPH_FINISHED: 'task.graph.finished' });
export function createTaskEvent(type, data = {}) {
    const event = /** @type {{ type: string, timestamp: string, [key: string]: any }} */ ({ type, timestamp: new Date().toISOString(), ...data });
    if (data.graphId !== undefined) event.graphId = data.graphId;
    if (data.taskId !== undefined) event.taskId = data.taskId;
    if (data.attemptId !== undefined) event.attemptId = data.attemptId;
    if (data.correlationId !== undefined) event.correlationId = data.correlationId;
    if (data.source !== undefined) event.source = Object.freeze({ ...data.source });
    if (data.target !== undefined) event.target = Object.freeze({ ...data.target });
    return event;
}
export function createEventSink(onEvent) { return (type, data) => { const event = createTaskEvent(type, data); if (typeof onEvent === 'function') onEvent(event); return event; }; }
