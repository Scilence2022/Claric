import { normalizeTask, normalizeCompound, TASK_STATES } from '../src/lib/task-runtime/task-model.js';
import { createResourceLocks } from '../src/lib/task-runtime/resource-locks.js';
import { executeTaskGraph, validateTaskGraph } from '../src/lib/task-runtime/task-graph.js';
import { TASK_EVENTS, createTaskEvent } from '../src/lib/task-runtime/task-events.js';
import { createProposalAggregate } from '../src/lib/task-runtime/proposal-aggregate.js';

test('task model normalizes compatibility fields and ids', () => {
  const task = normalizeTask({ id: 'legacy', instruction: '  do it  ' });
  expect(task.taskId).toBe('legacy');
  expect(task.instruction).toBe('do it');
  expect(task.attemptId).toBeTruthy();
  expect(normalizeCompound([task]).tasks).toHaveLength(1);
});

test('resource locks enforce ownership', () => {
  const locks = createResourceLocks();
  expect(locks.acquire(['word'], 'a')).toBe(true);
  expect(locks.canAcquire(['word'], 'b')).toBe(false);
  expect(locks.acquire(['word'], 'b')).toBe(false);
  locks.release(['word'], 'a');
  expect(locks.acquire(['word'], 'b')).toBe(true);
});

test('task graph validates dependencies and cycles', () => {
  expect(validateTaskGraph({ tasks: [{ taskId: 'a', dependsOn: ['missing'] }] }).valid).toBe(false);
  expect(validateTaskGraph({ tasks: [{ taskId: 'a', dependsOn: ['b'] }, { taskId: 'b', dependsOn: ['a'] }] }).valid).toBe(false);
});

test('task graph executes dependencies and emits lifecycle events', async () => {
  const events = [];
  const order = [];
  const result = await executeTaskGraph({ tasks: [{ taskId: 'a' }, { taskId: 'b', dependsOn: ['a'] }] }, async (task) => { order.push(task.taskId); return task.taskId; }, { onEvent: (event) => events.push(event) });
  expect(order).toEqual(['a', 'b']);
  expect(result.results.get('b').state).toBe(TASK_STATES.SUCCEEDED);
  expect(events.map((event) => event.type)).toContain(TASK_EVENTS.GRAPH_FINISHED);
});

test('task graph propagates abort and does not start later tasks', async () => {
  const controller = new AbortController();
  const started = [];
  await expect(executeTaskGraph({ tasks: [{ taskId: 'a' }, { taskId: 'b' }] }, async (task) => { started.push(task.taskId); controller.abort(); throw Object.assign(new Error('aborted'), { name: 'AbortError' }); }, { signal: controller.signal })).rejects.toThrow('aborted');
  expect(started).toEqual(['a']);
});

test('task events and proposal aggregate preserve metadata', () => {
  const event = createTaskEvent(TASK_EVENTS.TASK_STARTED, { taskId: 'a' });
  expect(event.type).toBe(TASK_EVENTS.TASK_STARTED);
  const aggregate = createProposalAggregate('g');
  expect(aggregate.add({ kind: 'format' }, { taskId: 'a', attemptId: '1' })).toMatchObject({ graphId: 'g', taskId: 'a', attemptId: '1' });
  expect(aggregate.size).toBe(1);
});
