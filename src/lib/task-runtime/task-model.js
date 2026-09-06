/** Task graph runtime data model and compatibility normalization. */

export const TASK_STATES = Object.freeze({ PENDING: 'pending', READY: 'ready', RUNNING: 'running', SUCCEEDED: 'succeeded', FAILED: 'failed', CANCELLED: 'cancelled', BLOCKED: 'blocked' });

let sequence = 0;
export function createId(prefix = 'task') { sequence += 1; return `${prefix}-${Date.now().toString(36)}-${sequence.toString(36)}`; }

export function normalizeTask(task = {}, index = 0) {
    const id = task.taskId || task.id || createId(`task-${index + 1}`);
    return {
        ...task,
        taskId: id,
        id,
        attemptId: task.attemptId || createId('attempt'),
        type: String(task.type || 'qa'),
        instruction: String(task.instruction || '').trim(),
        dependsOn: Array.isArray(task.dependsOn) ? [...new Set(task.dependsOn)] : [],
        resources: Array.isArray(task.resources) ? [...new Set(task.resources)] : [],
        state: task.state || TASK_STATES.PENDING,
    };
}

export function normalizeCompound(input = {}) {
    const tasks = Array.isArray(input) ? input : (input.tasks || input.nodes || []);
    return { ...(!Array.isArray(input) ? input : {}), graphId: input.graphId || createId('graph'), tasks: tasks.map(normalizeTask) };
}
