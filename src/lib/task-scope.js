/** Explicit task scope wins over an incidental live selection. */
export function requestedTaskScope(instruction = '') {
    const selection = /选区|所选|选中的?(?:部分|文本|段落|内容|表格)|选择的?(?:部分|文本|段落|内容)|\bselection\b|\bselected\s+(?:text|paragraphs?|passage|content|table)/i.test(instruction);
    const document = /全文|全篇|整篇|整个?文档|整[个篇]?文章|文档[里中]|\b(?:entire|whole)\s+(?:document|article)\b|\bdocument[- ]wide\b/i.test(instruction);
    return selection === document ? null : selection ? 'selection' : 'document';
}

export function resolveTaskScope(task, facts = {}, originalRequest = '') {
    if (['selection', 'document'].includes(task.scope)) return task.scope;
    return requestedTaskScope(task.instruction) || requestedTaskScope(originalRequest)
        || ((facts.hasSelection || facts.hasTextSelection || facts.hasImageSelection || facts.hasMultiCellTableRegion) ? 'selection' : 'document');
}
