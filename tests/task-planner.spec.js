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

  test('unknown task types are dropped with a warning', () => {
    const log = jest.fn();
    const tasks = parsePlan('[{"type":"delete","instruction":"删掉全文"},{"type":"format","instruction":"全文居中"}]', log);
    expect(tasks).toEqual([{ type: 'format', instruction: '全文居中' }]);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('unknown type'), 'warning');
  });

  test('tasks with empty instructions are dropped', () => {
    const log = jest.fn();
    const tasks = parsePlan('[{"type":"edit","instruction":"  "},{"type":"append","instruction":"续写结尾"}]', log);
    expect(tasks).toEqual([{ type: 'append', instruction: '续写结尾' }]);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('empty instruction'), 'warning');
  });

  test('non-object entries are dropped; non-string instructions are coerced', () => {
    const tasks = parsePlan('[null,"x",{"type":"edit","instruction":42}]');
    expect(tasks).toEqual([{ type: 'edit', instruction: '42' }]);
  });

  test('the plan is capped at MAX_TASKS with a warning', () => {
    const log = jest.fn();
    const many = Array.from({ length: 8 }, (_, i) => ({ type: 'format', instruction: `task ${i}` }));
    const tasks = parsePlan(JSON.stringify(many), log);
    expect(tasks).toHaveLength(6);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('capped at 6'), 'warning');
  });

  test('overlong instructions are truncated with a warning', () => {
    const log = jest.fn();
    const tasks = parsePlan(JSON.stringify([{ type: 'edit', instruction: 'x'.repeat(600) }]), log);
    expect(tasks[0].instruction).toHaveLength(500);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('truncated'), 'warning');
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
});
