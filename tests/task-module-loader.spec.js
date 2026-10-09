jest.mock('../src/lib/lazy-module-loader.js', () => ({ loadLazyModule: jest.fn(async () => ({})) }));
const { loadLazyModule } = require('../src/lib/lazy-module-loader.js');
const { preloadTaskModules, loadAgentActions, loadDocumentEditActions, loadFormatPlanning,
    loadTaskPlanner, loadTaskGraph, loadFileQuestion, loadCommentActions } = require('../src/taskpane/task-module-loader.js');

beforeEach(() => jest.clearAllMocks());

test('preflight loads all required code once before a composite request starts', async () => {
    const signal = new AbortController().signal;
    await preloadTaskModules([
        { type: 'format' }, { type: 'edit' }, { type: 'table_management' },
        { type: 'image_management' }, { type: 'document_edit' }, { type: 'comment_management' },
    ], { signal, visualFormatting: true });
    expect(loadLazyModule.mock.calls.map(([id]) => id).sort()).toEqual([
        'agent-actions', 'comment-actions', 'document-edit-actions', 'task-graph', 'visual-format',
    ]);
    for (const [, callback, options] of loadLazyModule.mock.calls) {
        expect(typeof callback).toBe('function');
        expect(options.signal).toBe(signal);
    }
});

test('plain body editing preflight does not download table or rendering tools', async () => {
    await preloadTaskModules([{ type: 'edit' }, { type: 'qa' }]);
    expect(loadLazyModule.mock.calls.map(([id]) => id)).toEqual(['task-graph']);
});

test('formatting on a host without visual tools does not require optional rendering code', async () => {
    await preloadTaskModules([{ type: 'format' }], { visualFormatting: false });
    expect(loadLazyModule.mock.calls.map(([id]) => id)).toEqual(['task-graph']);
});

test('failed capability code preflight rejects without running native actions', async () => {
    const error = new Error('Table module download failed');
    loadLazyModule.mockImplementationOnce(async () => ({})).mockRejectedValueOnce(error);
    await expect(preloadTaskModules([{ type: 'table_management' }])).rejects.toBe(error);
});

test.each([
    [loadAgentActions, 'agent-actions'], [loadDocumentEditActions, 'document-edit-actions'],
    [loadFormatPlanning, 'visual-format'], [loadTaskPlanner, 'task-planner'],
    [loadTaskGraph, 'task-graph'], [loadFileQuestion, 'file-question'], [loadCommentActions, 'comment-actions'],
])('module wrapper %p uses the shared loader with a cancellable import', async (load, id) => {
    const signal = new AbortController().signal;
    await load({ signal });
    expect(loadLazyModule).toHaveBeenCalledWith(id, expect.any(Function), expect.objectContaining({ signal, label: expect.any(String) }));
});
