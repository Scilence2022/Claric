/** Structured task lifecycle events. */
export const TASK_EVENTS = Object.freeze({ GRAPH_STARTED: 'task.graph.started', TASK_READY: 'task.ready', TASK_STARTED: 'task.started', TASK_SUCCEEDED: 'task.succeeded', TASK_FAILED: 'task.failed', GRAPH_FINISHED: 'task.graph.finished' });
export function createTaskEvent(type, data = {}) { return { type, timestamp: new Date().toISOString(), ...data }; }
export function createEventSink(onEvent) { return (type, data) => { const event = createTaskEvent(type, data); if (typeof onEvent === 'function') onEvent(event); return event; }; }
