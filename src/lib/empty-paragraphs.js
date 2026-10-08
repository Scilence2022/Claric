import { documentPartRoot } from './ooxml-text.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const BLANK = '(?:空\\s*段落|空白\\s*段落|空行|空白行)';
const CLEANUP_RE = new RegExp(
    '(?:删除|清除|清理|去掉|移除|去除).{0,24}' + BLANK
    + '|(?:多余|冗余|不必要)的?' + BLANK
    + '|\\b(?:delete|remove|clean\\s*up|get rid of|strip)\\b.{0,30}\\b(?:empty|blank|whitespace)\\b.{0,4}\\b(?:paragraphs?|lines?)'
    + '|\\b(?:extra|excess|redundant|unnecessary)\\s+(?:blank|empty)\\s+(?:paragraphs?|lines?)', 'i');
const KEEP_RE = new RegExp('(?:不要|勿|别|不必|无需|不需要).{0,4}(?:删除|清除|清理|去掉|移除|去除|改动|修改|调整|处理|动).{0,8}' + BLANK
    + '|(?:保留|保持).{0,8}' + BLANK
    + '|\\b(?:do not|don\u2019t|don\u0027t|never)\\s+(?:delete|remove|clean\\s*up|change|touch|alter|adjust)\\b.{0,20}\\b(?:blank|empty|whitespace)\\s+(?:paragraphs?|lines?)'
    + '|\\b(?:keep|preserve|retain|leave)\\b.{0,15}\\b(?:blank|empty|whitespace)\\b.{0,4}\\b(?:paragraphs?|lines?)'
    + '|\\bwithout\\s+(?:deleting|removing|changing)\\b.{0,20}\\b(?:blank|empty|whitespace)\\s+(?:paragraphs?|lines?)', 'i');

/** Only explicit blank-line cleanup authorizes paragraph deletion. */
export function requestsEmptyParagraphCleanup(instruction) {
    return typeof instruction === 'string' && !KEEP_RE.test(instruction) && CLEANUP_RE.test(instruction);
}

const SAFE_ELEMENTS = new Set([
    'p', 'pPr', 'r', 'rPr', 't', 'tab', 'proofErr', 'lastRenderedPageBreak',
    'pStyle', 'spacing', 'ind', 'jc', 'contextualSpacing', 'keepNext', 'keepLines',
    'widowControl', 'outlineLvl', 'textAlignment', 'tabs', 'adjustRightInd',
    'snapToGrid', 'suppressAutoHyphens', 'bidi', 'suppressLineNumbers',
    'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'color', 'sz', 'szCs', 'u',
    'strike', 'dstrike', 'caps', 'smallCaps', 'highlight', 'vertAlign', 'lang',
    'noProof', 'kern', 'position', 'w', 'rtl', 'cs', 'vanish', 'webHidden',
]);

/**
 * Empty text alone is insufficient: preserve fields, drawings, references,
 * breaks, sections, content controls, bookmarks and existing revisions.
 * Unknown XML stays read-only. A missing/invalid OOXML read is not evidence.
 */
export function isDeletableEmptyParagraphXml(ooxml) {
    const root = documentPartRoot(ooxml);
    if (!root) return false;
    const paragraphs = root.namespaceURI === W_NS && root.localName === 'p'
        ? [root] : Array.from(root.getElementsByTagNameNS(W_NS, 'p'));
    if (paragraphs.length !== 1) return false;
    const paragraph = paragraphs[0];
    for (let parent = paragraph.parentElement; parent; parent = parent.parentElement) {
        if (parent.namespaceURI === W_NS && !['body', 'document'].includes(parent.localName)) return false;
    }
    const elements = [paragraph, ...Array.from(paragraph.getElementsByTagName('*'))];
    return elements.every((element) => element.namespaceURI === W_NS && SAFE_ELEMENTS.has(element.localName)
        && (element.localName !== 't' || !element.textContent.trim()));
}
