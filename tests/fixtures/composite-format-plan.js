export const COMPOSITE_FORMAT_REQUEST = '全文优化格式，例如不正确的字体加粗，多余的空格，表格修改为三线格';

export function compositeFormatPlan() {
    return {
        requirements: [
            { id: 'r1', kind: 'action', outcome: 'document', text: '优化格式' },
            { id: 'r2', kind: 'action', outcome: 'document', text: '纠正不正确的字体加粗' },
            { id: 'r3', kind: 'action', outcome: 'document', text: '清理多余的空格' },
            { id: 'r4', kind: 'action', outcome: 'document', text: '表格修改为三线格' },
            { id: 'r5', kind: 'constraint', text: '处理范围为全文' },
        ],
        tasks: [
            { taskId: 'format', type: 'format', scope: 'document', instruction: '优化全文格式，结合结构纠正不正确的加粗，保留文字和语义。',
                covers: ['r1', 'r2', 'r5'], dependsOn: [] },
            { taskId: 'body', type: 'edit', scope: 'document', instruction: '清理全文正文的多余空格，保留有意义的空格、原有内容和结构。前一提案应用后重新读取文档。',
                covers: ['r3', 'r5'], dependsOn: ['format'] },
            { taskId: 'tables', type: 'table_management', scope: 'document', instruction: '检查全文所有原生表格，清理单元格多余空格，改为三线表：保留顶线、表头底线、底线，移除其他横线和竖线。前一提案应用后重新读取表格。',
                covers: ['r3', 'r4', 'r5'], dependsOn: ['body'] },
        ], unsupported: [],
    };
}

export function approvingFormatReview() {
    return { complete: true, unsupportedAccurate: true, checks: compositeFormatPlan().requirements
        .map(({ id }) => ({ requirementId: id, represented: true })), missing: [], invented: [], summary: 'All requested outcomes covered across native capabilities.' };
}
