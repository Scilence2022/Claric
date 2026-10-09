/** Native paragraph evidence and semantic selectors for formatting proposals. */
import { documentPartRoot, extractFinalTextFromOoxml } from './ooxml-text.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const key = (value) => String(value || '').toLowerCase().replace(/[\s_-]+/g, '');
const BODY_STYLES = new Set(['normal', 'nospacing', 'bodytext', 'bodytext2', 'bodytext3',
    'bodytextindent', 'bodytextindent2', 'bodytextindent3', 'bodytextfirstindent', 'bodytextfirstindent2']);
const NON_BODY_STYLES = /^(?:toc\d*|tocheading|bibliography|footnote(?:text|reference)|endnote(?:text|reference)|header|footer|listparagraph|quote|intensequote)$/;
const SECTION_NAME = /^(?:abstract|introduction|background|methods?|materials(?: and methods)?|results?|discussion|conclusions?|references|acknowledg(?:e)?ments?|appendi(?:x|ces)|摘要|引言|绪论|背景|材料(?:与|和)方法|方法|结果|讨论|结论|参考文献|致谢|附录)$/i;
const CAPTION = /^(?:(?:figure|fig\.?|table|scheme|plate)\s*[a-z]?\d+|(?:附?图|附?表)\s*[A-Za-z]?\d+)(?:[\s.:：、\-–—]|$)/i;
const NUMBERED_HEADING = /^(?:(?:\d+(?:[.\s]\d+)*\.?|[IVXLC]+\.)\s+\S|第[一二三四五六七八九十百\d]+[章节部分])/i;
const PROTECTED_MARKUP = new Set(['sectPr', 'sdt', 'del', 'moveFrom', 'moveTo', 'altChunk']);
const OBJECT_MARKUP = new Set(['drawing', 'pict', 'object', 'txbxContent']);
const NON_BODY_TARGETS = /标题|题注|表格|图表|图片|目录|参考文献|页眉|页脚|\b(?:headings?|titles?|captions?|tables?|images?|figures?|footnotes?|headers?|footers?|bibliograph(?:y|ies))\b/i;

function singleParagraph(ooxml) {
    const root = documentPartRoot(ooxml);
    if (!root) return null;
    const paragraphs = root.namespaceURI === W_NS && root.localName === 'p'
        ? [root] : Array.from(root.getElementsByTagNameNS(W_NS, 'p'));
    return paragraphs.length === 1 ? paragraphs[0] : null;
}
function on(element) { return !['0', 'false', 'off'].includes(element?.getAttributeNS(W_NS, 'val')); }
function sentenceLike(text) {
    return /[.!?。！？](?:[\s\u201d\u2019"')\]]|$)/u.test(text) && (text.length >= 35 || /[\u3400-\u9fff]/u.test(text) && text.length >= 15);
}

/**
 * A role is evidence, not a style-name synonym. Keep exact native text in the
 * descriptor; an unreadable/partial native paragraph is never assumed prose.
 * eligible is specifically eligibility for a semantic body-text selector.
 * @param {{index?: number, text?: string, style?: string, styleBuiltIn?: string,
 * ooxml?: string, inTable?: boolean, withinScope?: boolean, bold?: boolean,
 * outlineLevel?: number}} [input]
 */
export function describeFormatParagraph({ index, text, style, styleBuiltIn, ooxml, inTable, withinScope = true,
    bold, outlineLevel } = {}) {
    if (!Number.isInteger(index) || index < 0) throw new Error('Formatting paragraph index must be a non-negative integer.');
    const descriptor = { id: `p${index + 1}`, index, text: typeof text === 'string' ? text : '',
        style: typeof style === 'string' ? style : '', styleBuiltIn: typeof styleBuiltIn === 'string' ? styleBuiltIn : '',
        role: 'unknown', reason: 'Paragraph evidence is incomplete.', eligible: false, verified: false };
    const result = (role, reason, verified = descriptor.verified) => ({ ...descriptor, role, reason, verified,
        eligible: verified && role === 'body' });
    if (withinScope !== true) return result('protected', 'Paragraph is not fully inside the captured scope.');
    if (typeof text !== 'string') return result('unknown', 'Native paragraph text is unreadable.');
    const paragraph = singleParagraph(ooxml);
    if (!paragraph) return result('unknown', 'XML does not identify one readable paragraph.');
    if (typeof inTable !== 'boolean') return result('unknown', 'Native table membership is unreadable.');
    descriptor.verified = true;
    if (inTable) return result('table', 'Paragraph is inside a native Word table.');
    for (let parent = paragraph.parentElement; parent; parent = parent.parentElement) {
        if (parent.namespaceURI === W_NS && !['document', 'body'].includes(parent.localName)) {
            return result(parent.localName === 'tc' ? 'table' : 'protected', 'Paragraph has a protected native container.');
        }
    }
    const elements = [paragraph, ...Array.from(paragraph.getElementsByTagName('*'))];
    if (elements.some((element) => element.namespaceURI === W_NS && OBJECT_MARKUP.has(element.localName))) {
        return result('object', 'Paragraph contains a drawing, picture or embedded object.');
    }
    if (elements.some((element) => element.namespaceURI !== W_NS
        || PROTECTED_MARKUP.has(element.localName))) return result('protected', 'Paragraph contains protected or unsupported structure.');
    // Paragraph-level XML may include a range's synthetic trailing marker.
    // Only the exact native paragraph text, apart from its terminal mark, is compared.
    const xmlText = extractFinalTextFromOoxml(ooxml);
    if (xmlText !== text.replace(/\r$/, '')) return result('unknown', 'Native text and paragraph XML disagree.', false);
    const trimmed = text.trim();
    if (!trimmed) return result('empty', 'Paragraph has no prose text.');
    const builtIn = key(styleBuiltIn);
    if (/^heading[1-9]$/.test(builtIn)) return result('heading', 'Native built-in heading style.');
    if (['title', 'subtitle'].includes(builtIn)) return result('title', 'Native built-in title style.');
    if (builtIn === 'caption' || CAPTION.test(trimmed)) return result('caption', 'Native caption style or explicit caption label.');
    const outline = elements.find((element) => element.localName === 'outlineLvl');
    if ((outline && Number(outline.getAttributeNS(W_NS, 'val')) < 9)
        || Number.isInteger(outlineLevel) && outlineLevel >= 0 && outlineLevel < 9) return result('heading', 'Native paragraph has a heading outline level.');
    const short = trimmed.length <= 180;
    if (short && !sentenceLike(trimmed) && (NUMBERED_HEADING.test(trimmed) || SECTION_NAME.test(trimmed))) return result('heading', 'Text identifies a short numbered or named heading.');
    const keepNext = elements.some((element) => element.localName === 'keepNext' && on(element));
    const textRuns = Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r'))
        .filter((run) => run.getElementsByTagNameNS(W_NS, 't').length);
    const allBold = textRuns.length > 0 && textRuns.every((run) => Array.from(run.getElementsByTagNameNS(W_NS, 'b')).some(on));
    if (short && (keepNext || allBold || bold === true)) {
        return sentenceLike(trimmed)
            ? result('unknown', 'Short prose has manual heading formatting; its body role is uncertain.')
            : result('heading', 'Short paragraph has manual heading formatting.');
    }
    if (NON_BODY_STYLES.test(builtIn)) return result('unknown', 'Native style identifies non-body content.');
    const alignment = elements.find((element) => element.localName === 'jc')?.getAttributeNS(W_NS, 'val');
    if (!sentenceLike(trimmed) || /(?:@|\bORCID\b|^\s*(?:keywords?|author|affiliation|关键词|作者|单位)\s*[:：])/i.test(trimmed)
        || ['center', 'right'].includes(alignment) && !sentenceLike(trimmed)) {
        return result('unknown', 'Text is too ambiguous to identify as body prose.');
    }
    if (BODY_STYLES.has(builtIn)) return result('body', 'Native body style and prose text agree.');
    const bodyName = /(?:body|prose|正文)/i.test(style);
    // Long complete prose is positive evidence for custom manuscript styles;
    // arbitrary short custom-style paragraphs remain explicitly unresolved.
    const completeProse = sentenceLike(trimmed) && (bodyName || trimmed.length >= 100);
    if ((!builtIn || builtIn === 'other') && completeProse) return result('body', 'Custom style has complete prose text and no protected structural markers.');
    return result('unknown', 'Style and text do not establish a body paragraph.');
}

/** Resolve a selector against captured evidence. No rejected selector widens scope. */
export function resolveFormatParagraphs(inventory, op, { bodyOnly = false } = {}) {
    if (!Array.isArray(inventory)) throw new Error('Formatting paragraph inventory is unavailable.');
    if (!op || typeof op !== 'object') throw new Error('Formatting operation is unavailable.');
    const selectors = ['match', 'paragraphStyle', 'paragraphRole', 'paragraphIds'].filter((name) => op[name] !== undefined);
    if (selectors.length > 1) throw new Error('Formatting operation has conflicting target selectors.');
    if (op.paragraphRole !== undefined && op.paragraphRole !== 'body') throw new Error('Unknown formatting paragraph role.');
    const byId = new Map(inventory.map((paragraph) => [paragraph.id, paragraph]));
    if (byId.size !== inventory.length) throw new Error('Formatting paragraph inventory has duplicate IDs.');
    let ids;
    if (op.paragraphIds !== undefined) {
        if (!Array.isArray(op.paragraphIds) || !op.paragraphIds.length || op.paragraphIds.length > 500
            || new Set(op.paragraphIds).size !== op.paragraphIds.length
            || op.paragraphIds.some((id) => typeof id !== 'string' || !/^p[1-9]\d*$/.test(id) || !byId.has(id))) {
            throw new Error('Formatting paragraph IDs do not identify captured paragraphs.');
        }
        ids = new Set(op.paragraphIds);
    }
    const targets = [];
    const exclusions = [];
    for (const paragraph of inventory) {
        let reason;
        if (!paragraph.verified || paragraph.role === 'protected') reason = paragraph.reason;
        else if ((bodyOnly || op.paragraphRole === 'body') && (!paragraph.eligible || paragraph.role !== 'body')) reason = paragraph.reason;
        else if (ids && !ids.has(paragraph.id)) reason = 'Not selected by paragraph ID.';
        else if (op.paragraphStyle && key(paragraph.styleBuiltIn) !== key(op.paragraphStyle)) reason = 'Native built-in style does not match.';
        else if (op.match && !paragraph.text.includes(op.match)) reason = 'Exact text does not match.';
        if (reason) exclusions.push({ id: paragraph.id, index: paragraph.index, role: paragraph.role, reason });
        else targets.push(paragraph);
    }
    return { targets, exclusions };
}

/** Body-only authorization is explicit, never inferred from the Normal style. */
export function requestsBodyFormatting(instruction) {
    if (typeof instruction !== 'string') return false;
    const positive = instruction
        .replace(/(?:不要|勿|别|无需|不必|不)(?:将|把)?\s*正文(?:段落|内容)?\s*(?:修改|改动|变更|调整|处理|动)/g, '')
        .replace(/(?:不要|勿|别|无需|不必|不)\s*(?:修改|改动|变更|调整|处理|动)\s*正文(?:段落|内容)?/g, '')
        .replace(/(?:保留|保持)\s*正文(?:段落|内容)?\s*(?:原样|不变)|正文(?:段落|内容)?\s*(?:保持原样|保持不变|不变)/g, '')
        .replace(/\b(?:do not|don['’]t|without)\s+(?:chang(?:e|ing)|alter(?:ing)?|format(?:ting)?|touch(?:ing)?)\s+(?:the\s+)?(?:body\s+(?:text|paragraphs?|prose)|main\s+(?:text|prose))\b/gi, '')
        .replace(/\b(?:leave|keep|preserve)\s+(?:the\s+)?(?:body\s+(?:text|paragraphs?|prose)|main\s+(?:text|prose))\s+(?:unchanged|untouched|as is)\b/gi, '');
    const bodyMention = /(?<!非)正文|\b(?:body\s+(?:text|paragraphs?|prose)|main\s+(?:text|prose))\b/i.test(positive);
    if (!bodyMention) return false;
    // Body scope is not a blanket restriction on a compound request that also
    // formats headings/tables. Remove only clear preservation tails, keeping
    // any tail that contains a subsequent positive formatting action.
    const targets = positive.replace(/(?:^|[，,；;]|\band\s+)\s*(?:不(?:修改|更改|改动|调整|处理|改变|动)|不要(?:修改|更改|调整)|保留|保持|do not\s+(?:change|modify|alter|format|touch)|don['’]t\s+(?:change|modify|alter|format|touch)|without\s+(?:changing|modifying|altering|formatting|touching)|leave|preserve|keep)\s*([^。；;\n]*)/gi,
        (whole, rest) => NON_BODY_TARGETS.test(rest)
            && !/(?:修改|更改|改动|调整|设置|优化|加粗|居中|转换|改为|改成)|\b(?:change|modify|alter|format|align|center|centre|justify|set|make)\b/i.test(rest) ? '' : whole);
    return !NON_BODY_TARGETS.test(targets);
}

/** Deterministic fast path only for a single unambiguous body-alignment request. */
export function explicitBodyAlignmentOp(instruction) {
    if (!requestsBodyFormatting(instruction)) return null;
    const text = instruction.trim().replace(/[。.!！]$/, '');
    // Keep compound formatting, rewriting and preservation instructions in the
    // normal planner. This fast path must represent the entire request.
    const chinese = /^(?:请|将|把|请将|请把)?\s*(?:全文)?正文(?:段落)?\s*(?:(?:修改|改|设置|调整|设|变更)(?:为|成)|采用|使用)?\s*(两端对齐|左对齐|右对齐|居中对齐|居中)$/;
    const english = /^(?:please\s+)?(?:(?:set|change|make|align)\s+)?(?:the\s+)?(?:body\s+(?:text|paragraphs?|prose)|main\s+(?:text|prose))\s+(?:(?:alignment\s+)?(?:to\s+)?)(justified|justify|left(?:[- ]aligned)?|right(?:[- ]aligned)?|cent(?:er|re)(?:d|[- ]aligned)?)$/i;
    const match = text.match(chinese) || text.match(english);
    if (!match) return null;
    const value = match[1].toLowerCase();
    const alignment = /两端|justif/.test(value) ? 'justified' : /左|left/.test(value) ? 'left' : /右|right/.test(value) ? 'right' : 'centered';
    return { paragraphRole: 'body', paragraph: { alignment } };
}
