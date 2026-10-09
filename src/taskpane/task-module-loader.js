import { loadLazyModule } from '../lib/lazy-module-loader.js';
import { canReadWordVisuals } from './word-render-tools.js';

// Retry imports, never the native actions exported by those modules.
export const loadAgentActions = (options) => loadLazyModule('agent-actions',
    () => import(/* webpackChunkName: "agent-actions" */ './agent-actions.js'), { label: 'table and image tools', ...options });
export const loadDocumentEditActions = (options) => loadLazyModule('document-edit-actions',
    () => import(/* webpackChunkName: "document-edit-actions" */ './document-edit-actions.js'), { label: 'document editing tools', ...options });
export const loadFormatPlanning = (options) => loadLazyModule('visual-format',
    () => import(/* webpackChunkName: "visual-format" */ './format-planning-session.js'), { label: 'formatting tools', ...options });
export const loadFormatActions = (options) => loadLazyModule('format-actions',
    () => import(/* webpackChunkName: "format-actions" */ './word-format-actions.js'), { label: 'native formatting tools', ...options });
export const loadTaskPlanner = (options) => loadLazyModule('task-planner',
    () => import(/* webpackChunkName: "task-planner" */ '../lib/task-planner.js'), { label: 'task planner', ...options });
export const loadTaskGraph = (options) => loadLazyModule('task-graph',
    () => import(/* webpackChunkName: "task-planner" */ '../lib/task-runtime/task-graph.js'), { label: 'task execution tools', ...options });
export const loadFileQuestion = (options) => loadLazyModule('file-question',
    () => import(/* webpackChunkName: "file-question" */ '../lib/file-question.js'), { label: 'reference reading tools', ...options });
export const loadCommentActions = (options) => loadLazyModule('comment-actions',
    () => import(/* webpackChunkName: "comment-actions" */ './comment-actions.js'), { label: 'comment tools', ...options });

/** Check the whole plan's required code before reading Word or calling task models. */
export async function preloadTaskModules(tasks, options = {}) {
    const types = new Set(tasks.map((task) => task.type));
    const loads = [loadTaskGraph(options)];
    if (types.has('table_management') || types.has('image_management')) loads.push(loadAgentActions(options));
    if (types.has('document_edit')) loads.push(loadDocumentEditActions(options));
    if (types.has('format') || types.has('insert')) {
        loads.push(loadFormatActions(options));
        if (options.visualFormatting ?? canReadWordVisuals()) loads.push(loadFormatPlanning(options));
    }
    if (types.has('comment_management')) loads.push(loadCommentActions(options));
    await Promise.all(loads);
}
