/** Task graph runtime data model and compatibility normalization. */

export const TASK_STATES = Object.freeze({ PENDING: 'pending', READY: 'ready', RUNNING: 'running', SUCCEEDED: 'succeeded', FAILED: 'failed', CANCELLED: 'cancelled', BLOCKED: 'blocked' });

const TARGET_FIELDS = Object.freeze(['workspaceId', 'documentId', 'instanceId']);

let sequence = 0;
export function createId(prefix = 'task') { sequence += 1; return `${prefix}-${Date.now().toString(36)}-${sequence.toString(36)}`; }

export function normalizeTarget(target, defaults = {}) {
    if (typeof target === 'string') target = { documentId: target };
    if (!target || typeof target !== 'object' || Array.isArray(target)) return null;
    const normalized = {};
    for (const field of TARGET_FIELDS) {
        const value = target[field] ?? defaults[field];
        if (typeof value === 'string' && value.trim()) normalized[field] = value.trim();
    }
    return Object.keys(normalized).length ? Object.freeze(normalized) : null;
}

export function targetKey(target) {
    const normalized = normalizeTarget(target);
    return normalized ? TARGET_FIELDS.map((field) => normalized[field] || '').join(':') : '';
}

export function normalizeTask(task = {}, index = 0) {
    const id = task.taskId || task.id || createId(`task-${index + 1}`);
    const target = normalizeTarget(task.target || task.targetDocument || task.targetInstance);
    return {
        ...task,
        ...(target ? { target } : {}),
        taskId: id,
        id,
        attemptId: task.attemptId || createId('attempt'),
        type: String(task.type || 'qa'),
        instruction: String(task.instruction || '').trim(),
        dependsOn: Array.isArray(task.dependsOn) ? [...new Set(task.dependsOn)] : [],
        resources: Array.isArray(task.resources) ? [...new Set(task.resources)] : [],
        inputRefs: Array.isArray(task.inputRefs) ? task.inputRefs.map((ref) => typeof ref === 'string' ? ref : ref?.artifactId).filter(Boolean) : [],
        expectedRevision: typeof task.expectedRevision === 'string' ? task.expectedRevision : null,
        reviewRequired: task.reviewRequired !== false,
        state: task.state || TASK_STATES.PENDING,
    };
}

export function normalizeCompound(input = {}) {
    const tasks = Array.isArray(input) ? input : (input.tasks || input.nodes || []);
    return { ...(!Array.isArray(input) ? input : {}), graphId: input.graphId || createId('graph'), tasks: tasks.map(normalizeTask) };
}
