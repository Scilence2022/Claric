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

test('explicit runner failure blocks dependent tasks and still runs independent work', async () => {
  const execute = jest.fn(async (task) => task.taskId === 'a' ? { status: 'failed', error: new Error('Preparation failed') } : { status: 'answered' });
  const events = [];
  const { results } = await executeTaskGraph({ tasks: [
    { taskId: 'c', dependsOn: ['b'] }, { taskId: 'a' }, { taskId: 'b', dependsOn: ['a'] }, { taskId: 'd' },
  ] }, execute, { onEvent: (e) => events.push(e) });
  expect(execute.mock.calls.map(([t]) => t.taskId)).toEqual(['a', 'd']);
  expect(results.get('a').state).toBe('failed');
  expect(results.get('b').state).toBe('blocked');
  expect(results.get('c').state).toBe('blocked');
  expect(events.some((e) => e.taskId === 'a' && e.type === 'task.succeeded')).toBe(false);
});

test('staged results carry artifacts into later tasks but emit staged rather than applied success', async () => {
  const events = [];
  const execute = jest.fn(async (task, ctx) => task.taskId === 'a'
    ? { status: 'staged', artifacts: [{ after: 'Draft text' }] }
    : { status: 'answered', summary: ctx.inputs[1].value.after });
  const result = await executeTaskGraph({ tasks: [{ taskId: 'a' }, { taskId: 'b', dependsOn: ['a'], inputRefs: ['a:0'] }] }, execute, { onEvent: (e) => events.push(e) });
  expect(result.results.get('b').value.summary).toBe('Draft text');
  expect(events.some((e) => e.type === 'task.staged')).toBe(true);
});

test('a later graph pass reuses completed work and retries only blocked descendants', async () => {
  const graph = { tasks: [{ taskId: 'a' }, { taskId: 'b', dependsOn: ['a'], inputRefs: ['a:0'] }, { taskId: 'c' }] };
  const execute = jest.fn(async (task, ctx) => {
    if (task.taskId === 'a') return { status: 'staged', artifacts: [{ text: 'new content' }] };
    if (task.taskId === 'b') return ctx.inputs[0].value.status === 'staged'
      ? { status: 'blocked', error: new Error('Waiting for Word application') }
      : { status: 'staged', summary: ctx.inputs[1].value.text };
    return { status: 'answered', summary: 'independent' };
  });
  const first = await executeTaskGraph(graph, execute);
  expect(first.results.get('b').state).toBe('blocked');
  const retained = new Map([...first.results].filter(([, result]) => result.state === 'succeeded'));
  retained.set('a', { ...retained.get('a'), value: { ...retained.get('a').value, status: 'applied' } });
  const second = await executeTaskGraph(graph, execute, { initialResults: retained });
  expect(execute.mock.calls.map(([task]) => task.taskId)).toEqual(['a', 'c', 'b', 'b']);
  expect(second.results.get('b').value.summary).toBe('new content');
  expect(second.results.get('a').value.status).toBe('applied');
});

test.each([false, true])('only verified no-op satisfies dependencies: %s', async (satisfied) => {
  const execute = jest.fn(async () => ({ status: 'no_op', satisfied }));
  const { results } = await executeTaskGraph({ tasks: [{ taskId: 'a' }, { taskId: 'b', dependsOn: ['a'] }] }, execute);
  expect(execute).toHaveBeenCalledTimes(satisfied ? 2 : 1);
  expect(results.get('a').state).toBe(satisfied ? 'succeeded' : 'failed');
  if (!satisfied) expect(results.get('b').state).toBe('blocked');
});

test('missing artifacts and explicit blocked outcomes do not report success', async () => {
  const execute = jest.fn(async () => ({ status: 'blocked', error: new Error('Awaiting application') }));
  const { results } = await executeTaskGraph({ tasks: [{ taskId: 'a', inputRefs: ['missing'] }, { taskId: 'b' }] }, execute);
  expect(execute).toHaveBeenCalledTimes(1);
  expect([...results.values()].map((r) => r.state)).toEqual(['blocked', 'blocked']);
});

test('abort after a normal runner return stops the rest of the ready batch', async () => {
  const controller = new AbortController();
  const execute = jest.fn(async () => { controller.abort(); return { status: 'staged' }; });
  await expect(executeTaskGraph({ tasks: [{ taskId: 'a' }, { taskId: 'b' }] }, execute, { signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
  expect(execute).toHaveBeenCalledTimes(1);
});
