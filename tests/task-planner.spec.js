/**
 * Task planner tests: the JSON contract between the planner LLM and the
 * compound-turn executor. parsePlan must tolerate fences/prose and enforce
 * the task-type allowlist so only known pipelines get dispatched.
 */

const { buildPlanPrompt, parsePlan } = require('../src/lib/task-planner.js');

describe('parsePlan', () => {
  test('parses a bare JSON task array', () => {
    const tasks = parsePlan('[{"type":"insert","instruction":"增加标题"},{"type":"edit","instruction":"深度润色"}]');
    expect(tasks).toEqual([
      { type: 'insert', instruction: '增加标题' },
      { type: 'edit', instruction: '深度润色' },
    ]);
  });

  test('accepts the table task type', () => {
    const tasks = parsePlan('[{"type":"table","instruction":"在文档末尾插入一个三行三列的表格"}]');
    expect(tasks).toEqual([{ type: 'table', instruction: '在文档末尾插入一个三行三列的表格' }]);
  });

  test('accepts the document-scope image/table management task types', () => {
    const tasks = parsePlan('[{"type":"image_management","instruction":"给所有图片加上标题"},{"type":"table_management","instruction":"把表格改成三线表样式"}]');
    expect(tasks).toEqual([
      { type: 'image_management', instruction: '给所有图片加上标题' },
      { type: 'table_management', instruction: '把表格改成三线表样式' },
    ]);
  });

  test('strips code fences and tolerates surrounding prose', () => {
    const tasks = parsePlan('Sure:\n```json\n[{"type":"qa","instruction":"总结全文"}]\n```\nDone.');
    expect(tasks).toEqual([{ type: 'qa', instruction: '总结全文' }]);
  });

  test('invalid JSON returns null with a warning', () => {
    const log = jest.fn();
    expect(parsePlan('[{oops}]', log)).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('not valid JSON'), 'warning');
  });

  test('no JSON array returns null with a warning', () => {
    const log = jest.fn();
    expect(parsePlan('I cannot decompose that.', log)).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('no JSON array'), 'warning');
  });

  test('empty / nullish input returns null', () => {
    expect(parsePlan('')).toBeNull();
    expect(parsePlan(null)).toBeNull();
    expect(parsePlan(undefined)).toBeNull();
  });

  test('an empty task array returns null (caller falls back)', () => {
    const log = jest.fn();
    expect(parsePlan('[]', log)).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('no valid tasks'), 'warning');
  });

  test('unknown task types reject the entire plan rather than dropping a requirement', () => {
    const log = jest.fn();
    const tasks = parsePlan('[{"type":"delete","instruction":"删掉全文"},{"type":"format","instruction":"全文居中"}]', log);
    expect(tasks).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('unknown type'), 'warning');
  });

  test('tasks with empty instructions reject the plan', () => {
    const log = jest.fn();
    const tasks = parsePlan('[{"type":"edit","instruction":"  "},{"type":"append","instruction":"续写结尾"}]', log);
    expect(tasks).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('empty instruction'), 'warning');
  });

  test('malformed entries reject the plan without coercing instructions', () => {
    const tasks = parsePlan('[null,"x",{"type":"edit","instruction":42}]');
    expect(tasks).toBeNull();
  });

  test('too many tasks reject the plan without truncating it', () => {
    const log = jest.fn();
    const many = Array.from({ length: 8 }, (_, i) => ({ type: 'format', instruction: `task ${i}` }));
    const tasks = parsePlan(JSON.stringify(many), log);
    expect(tasks).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('exceeds 6'), 'warning');
  });

  test('long instructions retain constraints and over-limit instructions are rejected', () => {
    const log = jest.fn();
    const tasks = parsePlan(JSON.stringify([{ type: 'edit', instruction: 'x'.repeat(600) }]), log);
    expect(tasks[0].instruction).toHaveLength(600);
    expect(parsePlan(JSON.stringify([{ type: 'edit', instruction: 'x'.repeat(4001) }]), log)).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('exceeds 4000'), 'warning');
  });
});

describe('buildPlanPrompt', () => {
  test('embeds the instruction and lists every capability', () => {
    const p = buildPlanPrompt('增加标题，并深度润色修改', false);
    expect(p).toContain('增加标题，并深度润色修改');
    for (const type of ['"insert"', '"format"', '"edit"', '"append"', '"table"', '"illustration"', '"qa"']) {
      expect(p).toContain(type);
    }
    // Document-scope image/table management are planner-recognizable too.
    for (const type of ['"image_management"', '"table_management"']) {
      expect(p).toContain(type);
    }
    expect(p).toContain('Output ONLY a JSON array');
  });

  test('states the selection context both ways', () => {
    expect(buildPlanPrompt('x', true)).toContain('has a text selection');
    expect(buildPlanPrompt('x', false)).toContain('NO text selection');
  });

  test('the OUTPUT CONTRACT enum accepts every TASK_TYPE the parser allows', () => {
    // Regression: the contract enum used to omit image_management /
    // table_management while the CAPABILITIES section taught them, so a
    // contract-obeying model could never emit those two task types and the
    // corresponding compound sub-tasks misrouted.
    const p = buildPlanPrompt('x', false);
    const contract = p.slice(p.indexOf('OUTPUT CONTRACT'));
    expect(contract).toContain(
      'insert|format|edit|append|table|illustration|qa|image_management|table_management'
    );
    // And no type known to the parser is missing from the enum.
    for (const type of [
      'insert', 'format', 'edit', 'append', 'table', 'illustration', 'qa',
      'image_management', 'table_management',
    ]) {
      expect(contract).toContain(type);
    }
  });
});

test('document editing plans preserve dependency, resource and artifact references', () => {
  expect(parsePlan(JSON.stringify([
    { taskId: 'draft', type: 'document_edit', instruction: 'Integrate discussion', resources: ['body'] },
    { id: 'explain', type: 'qa', instruction: 'Explain the proposal', dependsOn: ['draft'], inputRefs: ['draft:0'] },
  ]))).toEqual([
    { taskId: 'draft', type: 'document_edit', instruction: 'Integrate discussion', resources: ['body'] },
    { taskId: 'explain', type: 'qa', instruction: 'Explain the proposal', dependsOn: ['draft'], inputRefs: ['draft:0'] },
  ]);
});

test.each([
  [{ taskId: 'a', type: 'edit', instruction: 'x', dependsOn: ['missing'] }],
  [{ taskId: 'a', type: 'edit', instruction: 'x', dependsOn: ['a'] }],
  [{ taskId: 'a', type: 'edit', instruction: 'x' }, { taskId: 'a', type: 'qa', instruction: 'y' }],
  [{ taskId: 'bad id', type: 'edit', instruction: 'x' }],
  [{ type: 'edit', instruction: 'x', resources: 'body' }],
  [{ type: 'edit', instruction: 'x', inputRefs: [42] }],
])('invalid graph metadata rejects the complete plan', (tasks) => {
  expect(parsePlan(JSON.stringify(tasks))).toBeNull();
});
