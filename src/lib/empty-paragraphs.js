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
    'outline', 'shadow', 'emboss', 'imprint', 'effect', 'fitText',
    'suppressPunctuationKerning', 'kinsoku', 'wordWrap', 'overflowPunct',
    'autoSpaceDE', 'autoSpaceDN', 'textDirection', 'mirrorIndents',
    // Formatting history contains properties, not deleted/inserted prose.
    // Word's native paragraph deletion handles these tracked revisions.
    'pPrChange', 'rPrChange',
    'noProof', 'kern', 'position', 'w', 'rtl', 'cs', 'vanish', 'webHidden',
]);

/**
 * Empty text alone is insufficient: preserve fields, drawings, references,
 * page/column breaks, sections, content controls, bookmarks and deleted/moved revisions.
 * Ordinary formatting history and inserted whitespace are safe only when
 * every nested property and run is safe.
 * Unknown XML stays read-only. A missing/invalid OOXML read is not evidence.
 */
export function inspectEmptyParagraphXml(ooxml) {
    const blocked = (reason, element) => ({ deletable: false, reason, ...(element ? { element } : {}) });
    const root = documentPartRoot(ooxml);
    if (!root) return blocked('invalid XML');
    const paragraphs = root.namespaceURI === W_NS && root.localName === 'p'
        ? [root] : Array.from(root.getElementsByTagNameNS(W_NS, 'p'));
    if (paragraphs.length !== 1) return blocked('XML does not identify one paragraph');
    const paragraph = paragraphs[0];
    for (let parent = paragraph.parentElement; parent; parent = parent.parentElement) {
        if (parent.namespaceURI === W_NS && !['body', 'document'].includes(parent.localName)) return blocked('protected container');
    }
    const elements = [paragraph, ...Array.from(paragraph.getElementsByTagName('*'))];
    for (const element of elements) {
        if (element.namespaceURI !== W_NS) return blocked('unsupported XML namespace');
        const name = element.localName;
        if (name === 'del') {
            const parent = element.parentElement;
            const deletedMark = parent?.localName === 'rPr' && parent.parentElement?.localName === 'pPr'
                && parent.parentElement.parentElement === paragraph;
            return blocked(deletedMark ? 'paragraph mark already tracked as deleted' : 'deleted revision content', 'w:del');
        }
        // An inserted EMPTY paragraph mark or whitespace run is ordinary blank
        // content. Native tracked deletion handles its insertion history; never
        // accept/reject revisions or rewrite XML to make a target eligible.
        if (name === 'ins') {
            const parent = element.parentElement;
            const insertedMark = parent?.localName === 'rPr' && parent.parentElement?.localName === 'pPr'
                && !element.children.length && !element.textContent.trim();
            const insertedRuns = parent === paragraph && Array.from(element.children)
                .every((child) => child.namespaceURI === W_NS && child.localName === 'r')
                && Array.from(element.childNodes).every((child) => child.nodeType === 1 || !child.textContent.trim());
            if (insertedMark || insertedRuns) continue;
            return blocked('protected revision', 'w:ins');
        }
        // A disabled pagination flag does not carry a page break.
        if (name === 'pageBreakBefore' && element.parentElement?.localName === 'pPr'
            && ['0', 'false', 'off'].includes(element.getAttributeNS(W_NS, 'val'))
            && !element.children.length) continue;
        // Soft line breaks in an otherwise blank paragraph are blank lines,
        // unlike explicit page/column breaks or text-wrapping clear commands.
        if (['br', 'cr'].includes(name) && element.parentElement?.localName === 'r'
            && !element.children.length && !element.textContent.trim()) {
            const type = element.getAttributeNS(W_NS, 'type');
            const clear = element.getAttributeNS(W_NS, 'clear');
            if ((!type || type === 'textWrapping') && (!clear || clear === 'none')) continue;
        }
        if (!SAFE_ELEMENTS.has(name)) {
            // Schema names only: no text, XML, author, bookmark name or IDs.
            const label = /^[A-Za-z][A-Za-z0-9]{0,47}$/.test(name) ? `w:${name}` : undefined;
            return blocked('protected or unsupported markup', label);
        }
        if (name === 't' && element.textContent.trim()) return blocked('non-whitespace XML text');
    }
    return { deletable: true, reason: null };
}

export function isDeletableEmptyParagraphXml(ooxml) {
    return inspectEmptyParagraphXml(ooxml).deletable;
}
