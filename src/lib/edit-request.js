import { intentText as stripConstraintClauses } from './intent-text.js';

const ENGLISH_STRUCTURE = /\b(?:reorder|rearrange|reorgani[sz]e|restructure|reshuffle|swap|move)\s+(?:(?:the|this|these|all|selected|body|existing|whole|first|last|two|\d+)\s+){0,5}(?:paragraphs?|sections?|chapters?|passages?|article|paper|manuscript|document|introduction|background|methods|results|discussion|conclusion|abstract)\b/i;
const CHINESE_STRUCTURE = /(?:重排|重组|重构|重新(?:组织|排列|安排)|调整).{0,16}(?:(?:文章|论文|文档|全文|正文)(?:的)?结构|(?:段落|章节|部分)(?:的)?(?:顺序|次序))|(?:重排|重组|重构|重新排列).{0,6}(?:段落|章节|文章|论文|正文)|(?:段落|章节)(?:的)?(?:顺序|次序).{0,12}(?:调整|改为|重排)|(?:文章|论文|文档|全文|正文)(?:的)?结构.{0,12}(?:调整|重排|重组)|(?:调整|优化).{0,12}(?:行文|叙述|论证)(?:顺序|结构)/;
const CHINESE_MOVE = /(?:段落|章节|[一二三四五六七八九十\d]+[节章段]|摘要|引言|背景|方法|结果|讨论|结论).{0,28}(?:移到|移至|挪到|提前到|放到|放在|调到|对调|互换)|(?:移动|挪动|对调|互换).{0,12}(?:段落|章节|[节章段])/;

/** Cheap routing hints only. The document session resolves the actual goal and scope. */
export function inspectEditRequest(input) {
    let action = String(input || '').trim();
    const explanatory = /^(?:please\s+)?(?:how|why|what|explain|describe)\b|^(?:请)?(?:如何|怎么|怎样|为什么|解释|说明)/i.test(action);
    const polite = /^(?:please\s+)?(?:can|could|would|will)\s+you\s+(?:please\s+)?|^(?:请问[，,\s]*)?(?:你|您)?(?:能否|能不能|能|可以|可否)(?:帮我|帮忙)?/i;
    if (!explanatory && polite.test(action)) action = action.replace(polite, '').replace(/(?:吗|呢)?[?？]?\s*$/, '').trim();
    // Preserve the complete user request elsewhere; negated and preservation
    // clauses are constraints, not actions (shared with the intent detectors).
    const intentText = stripConstraintClauses(action);
    const actionVerb = /\b(?:insert|integrate|incorporate|weave|supplement|add|expand)\b|插入|补充|融入|整合|融合|增补|扩充/i.test(intentText);
    const prose = /\b(?:text|paragraphs?|passages?|discussion|content|article|paper|manuscript|sections?|document)\b|正文|段落|讨论|内容|文章|论文|章节|文档/i.test(intentText);
    const placement = /\b(?:appropriate|suitable|relevant|between|before|after)\b|合适|适当|相应|之间|之前|之后/i.test(intentText);
    const objectCreation = /\b(?:tables?|images?|pictures?|diagrams?|charts?|svg)\b|表格|图片|图像|插图|示意图|流程图/i.test(intentText);
    const explicitEnd = /\b(?:append|at the end|to the end|document end)\b|文末|文档末尾|追加|续写/i.test(intentText);
    const negative = /^(?:please\s+)?(?:do not|don't|never)\b|^(?:请)?(?:不要|别|禁止)/i.test(intentText);
    const question = /[?？]$|吗$/.test(action);
    const canEdit = !explanatory && !question && !negative;
    const needsStructureEdit = canEdit && (ENGLISH_STRUCTURE.test(intentText) || CHINESE_STRUCTURE.test(intentText) || CHINESE_MOVE.test(intentText));
    return { intentText, needsStructureEdit, needsDocumentEdit: needsStructureEdit
        || (canEdit && actionVerb && (prose || placement) && !objectCreation && !explicitEnd) };
}
