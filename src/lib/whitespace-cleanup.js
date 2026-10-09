/** Conservative contracts for requests to remove redundant spaces only. */
export function requestsWhitespaceOnlyCleanup(instruction = '') {
    const text = String(instruction);
    const cleanup = /(?:清理|去除|删除|移除|去掉|整理|消除).{0,30}(?:多余|冗余|重复).{0,8}空格|\b(?:remove|clean(?:\s*up)?|delete|trim|collapse)\b.{0,40}\b(?:extra|redundant|duplicate|unnecessary|excess)\s+spaces?\b/i.test(text);
    if (!cleanup) return false;
    // Negated preservation clauses are common in the planner's instructions.
    const actions = text.replace(/(?:不|勿|不得|不要|不做)[^，。；;,.!?\n]{0,12}?(?:改写|重写|润色|翻译|扩写|(?:修改|改动|调整|更改|改变)(?:内容|格式|字体|字号)|加粗|斜体|下划线)/g, '')
        .replace(/(?:保持|保留|维持)(?:原有|已有|当前|现有)?(?:字体格式|格式|字体|字号|加粗|粗体|斜体|下划线|样式)/g, '')
        .replace(/\b(?:do not|don't|without|never)\s+(?:(?:change|changing|alter|altering|modify|modifying|adjust|adjusting)\s+(?:the\s+)?(?:existing\s+)?)?(?:rewrite|rewriting|rephrase|bold|bolding|format|formatting|styles?|fonts?|italic(?:ize|izing)?|underline|underlining)\b/gi, '')
        .replace(/\b(?:keep|preserve|retain)\s+(?:the\s+)?(?:(?:existing|current|original)\s+)?(?:formatting|format|fonts?|styles?|bold(?:ing)?|italics?|underlining|alignment|indentation)\b/gi, '');
    return !/(?:润色|翻译|扩写|改写|重写|加粗|三线|字体|字号|标题样式|格式|样式|对齐|缩进|斜体|下划线|空行|空段落)|\b(?:polish|translate|rewrite|rephrase|bold|font|three.line|format(?:ting)?|italic(?:ize|izing|s)?|underline(?:d|ing)?|styles?|alignment|indentation|blank\s+(?:lines|paragraphs))\b/i.test(actions);
}

export function normalizeWhitespaceText(text = '') {
    return String(text).replace(/\r\n?/g, '\n');
}

/** A no-op needs positive evidence from the source, rather than a model echo. */
export function hasRedundantSpaceCandidates(text = '') {
    return normalizeWhitespaceText(text).split('\n').some((line) => /^ +| +$| {2,}| +[,.;!?]|[\u3400-\u9fff] +[\u3400-\u9fff，。！？；：、）】]|[（【] +[\u3400-\u9fff]/u.test(line));
}

/**
 * Only ASCII-space deletion is allowed. Newlines, words, punctuation, tabs and
 * non-breaking spaces survive exactly; a single word separator must survive.
 * Uncertain single-space deletions are rejected rather than joining words.
 */
export function validateWhitespaceCleanup(original, amendment) {
    const before = normalizeWhitespaceText(original).split('\n');
    const after = normalizeWhitespaceText(amendment).split('\n');
    if (before.length !== after.length) return { valid: false, reason: 'Space cleanup must preserve paragraph structure.' };
    for (let i = 0; i < before.length; i++) {
        if (before[i].replace(/ /g, '') !== after[i].replace(/ /g, '')) {
            return { valid: false, reason: 'Space cleanup changed content or protected whitespace.' };
        }
        const originalParts = before[i].split(/( +)/);
        const amendedParts = after[i].split(/( +)/);
        // Compare the space count at each position between non-space chars.
        const positions = (parts) => {
            let offset = 0;
            const spaces = new Map();
            for (const part of parts) {
                if (/^ +$/.test(part)) spaces.set(offset, part.length);
                else offset += part.length;
            }
            return spaces;
        };
        const oldSpaces = positions(originalParts);
        const newSpaces = positions(amendedParts);
        const content = before[i].replace(/ /g, '');
        const contentLength = content.length;
        for (const [offset, count] of newSpaces) {
            if (count > (oldSpaces.get(offset) || 0)) return { valid: false, reason: 'Space cleanup may remove spaces but must not insert them.' };
        }
        for (const [offset, count] of oldSpaces) {
            const nextCount = newSpaces.get(offset) || 0;
            if (count === nextCount || offset === 0 || offset === contentLength) continue;
            const left = content[offset - 1];
            const right = content[offset];
            const punctuationGap = /[,.;!?]/.test(right);
            const cjkGap = /[\u3400-\u9fff]/u.test(left) && /[\u3400-\u9fff，。！？；：、）】]/u.test(right)
                || /[（【]/u.test(left) && /[\u3400-\u9fff]/u.test(right);
            if (nextCount === 0 && !punctuationGap && !cjkGap) {
                return { valid: false, reason: 'Space cleanup must preserve meaningful internal separators.' };
            }
        }
    }
    return { valid: true };
}
