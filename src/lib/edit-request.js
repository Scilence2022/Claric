/** Cheap routing hints only. The document session resolves the actual goal and scope. */
export function inspectEditRequest(input) {
    let action = String(input || '').trim();
    const explanatory = /^(?:please\s+)?(?:how|why|what|explain|describe)\b|^(?:请)?(?:如何|怎么|怎样|为什么|解释|说明)/i.test(action);
    const polite = /^(?:please\s+)?(?:can|could|would|will)\s+you\s+(?:please\s+)?|^(?:请问[，,\s]*)?(?:你|您)?(?:能否|能不能|能|可以|可否)(?:帮我|帮忙)?/i;
    if (!explanatory && polite.test(action)) action = action.replace(polite, '').replace(/(?:吗|呢)?[?？]?\s*$/, '').trim();
    // Preserve the complete user request elsewhere; these clauses are not formatting actions.
    const intentText = action.replace(/(?:[，,;；]\s*|^)(?:保留|保持|不要改变|不要修改|不改变|preserve\b|keep\b)[^，,;；。]*(?:格式|标题|样式|formatting|formats?|headings?|styles?)[^，,;；。]*/gi, '').trim();
    const actionVerb = /\b(?:insert|integrate|incorporate|weave|supplement|add|expand)\b|插入|补充|融入|整合|融合|增补|扩充/i.test(intentText);
    const prose = /\b(?:text|paragraphs?|passages?|discussion|content|article|paper|manuscript|sections?|document)\b|正文|段落|讨论|内容|文章|论文|章节|文档/i.test(intentText);
    const placement = /\b(?:appropriate|suitable|relevant|between|before|after)\b|合适|适当|相应|之间|之前|之后/i.test(intentText);
    const objectCreation = /\b(?:tables?|images?|pictures?|diagrams?|charts?|svg)\b|表格|图片|图像|插图|示意图|流程图/i.test(intentText);
    const explicitEnd = /\b(?:append|at the end|to the end|document end)\b|文末|文档末尾|追加|续写/i.test(intentText);
    const negative = /^(?:please\s+)?(?:do not|don't|never)\b|^(?:请)?(?:不要|别|禁止)/i.test(intentText);
    const question = /[?？]$|吗$/.test(action);
    return { intentText, needsDocumentEdit: !explanatory && !question && !negative && actionVerb && (prose || placement)
        && !objectCreation && !explicitEnd };
}
