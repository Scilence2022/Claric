const { resolveTaskScope, requestedTaskScope } = require('../src/lib/task-scope.js');

test.each(['全文优化格式', '清理整个文档的多余空格', 'Format the entire document', 'Polish the whole article', 'document-wide formatting'])('explicit document scope overrides an incidental selection: %s', (instruction) => {
    expect(resolveTaskScope({ instruction }, { hasSelection: true })).toBe('document');
});

test('per-task scopes preserve mixed document and selection requests', () => {
    const request = '全文优化格式，但仅润色选中段落';
    expect(resolveTaskScope({ instruction: '优化格式', scope: 'document' }, { hasSelection: true }, request)).toBe('document');
    expect(resolveTaskScope({ instruction: '润色', scope: 'selection' }, { hasSelection: true }, request)).toBe('selection');
    expect(resolveTaskScope({ instruction: '润色所选段落' }, { hasSelection: true }, '全文优化格式')).toBe('selection');
});

test('unscoped text retains selection behavior and location choice is not selection scope', () => {
    expect(resolveTaskScope({ instruction: '修正加粗' }, { hasSelection: true })).toBe('selection');
    expect(resolveTaskScope({ instruction: '修正加粗' }, {}, '全文优化格式')).toBe('document');
    expect(requestedTaskScope('选择合适的位置插入讨论')).toBeNull();
});
