/** Read OMML as immutable inline text, with an exact Word native projection.
 * Word's text API separates structured equation arguments with CRs, omitting
 * operators such as fraction bars. Those CRs are not paragraph boundaries.
 * Never guess the native representation of an unknown equation construct. */
export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

const latin = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
const greek = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ∇αβγδεζηθικλμνξοπρςστυφχψω∂ϵϑϰϕϱϖ';
const starts = { b: 0x1d400, i: 0x1d434, bi: 0x1d468,
    script: 0x1d49c, 'script-b': 0x1d4d0, fraktur: 0x1d504,
    'double-struck': 0x1d538, 'fraktur-b': 0x1d56c,
    'sans-serif': 0x1d5a0, 'sans-serif-b': 0x1d5d4,
    'sans-serif-i': 0x1d608, 'sans-serif-bi': 0x1d63c, monospace: 0x1d670 };
const holes = { 0x1d455: 'ℎ', 0x1d49d: 'ℬ', 0x1d4a0: 'ℰ', 0x1d4a1: 'ℱ',
    0x1d4a3: 'ℋ', 0x1d4a4: 'ℐ', 0x1d4a7: 'ℒ', 0x1d4a8: 'ℳ', 0x1d4ad: 'ℛ',
    0x1d4ba: 'ℯ', 0x1d4bc: 'ℊ', 0x1d4c4: 'ℴ', 0x1d506: 'ℭ', 0x1d50b: 'ℌ',
    0x1d50c: 'ℑ', 0x1d515: 'ℜ', 0x1d51d: 'ℨ', 0x1d53a: 'ℂ', 0x1d53f: 'ℍ',
    0x1d545: 'ℕ', 0x1d547: 'ℙ', 0x1d548: 'ℚ', 0x1d549: 'ℝ', 0x1d551: 'ℤ' };

function attribute(node, name, fallback = '') {
    return node?.getAttributeNS(MATH_NS, name) ?? fallback;
}
function property(node, name) { return Array.from(node?.children || []).find((child) => child.localName === name); }

function runText(node) {
    const properties = property(node, 'rPr');
    const plain = property(properties, 'nor');
    const style = attribute(property(properties, 'sty'), 'val', 'i');
    const script = attribute(property(properties, 'scr'), 'val', 'roman');
    const key = script === 'roman' ? style : script === 'sans-serif'
        ? script + (style === 'p' ? '' : '-' + style) : script + (style.includes('b') ? '-b' : '');
    const start = starts[key];
    let text = '';
    for (const child of Array.from(node.children)) {
        if (child.localName !== 't' || child.namespaceURI !== MATH_NS) continue;
        for (const char of child.textContent || '') {
            if (plain) { text += char; continue; }
            const index = latin.indexOf(char);
            const greekIndex = greek.indexOf(char);
            const greekStart = { b: 0x1d6a8, i: 0x1d6e2, bi: 0x1d71c }[style];
            const digitStart = { b: 0x1d7ce, 'double-struck': 0x1d7d8, 'sans-serif': 0x1d7e2,
                'sans-serif-b': 0x1d7ec, monospace: 0x1d7f6 }[key];
            if (index >= 0 && start) text += holes[start + index] || String.fromCodePoint(start + index);
            else if (greekIndex >= 0 && script === 'roman' && greekStart) text += String.fromCodePoint(greekStart + greekIndex);
            else if (digitStart && /[0-9]/.test(char)) text += String.fromCodePoint(digitStart + Number(char));
            else text += char === '*' ? '∗' : char;
        }
    }
    return text;
}

/** Canonical text retains the equation's operators; nativeText matches the
 * host's linear text exactly, including nested argument boundary controls.
 * reviewed=true selects Mac getReviewedText's punctuation-delimited format,
 * only for structures confirmed by captured native/XML pairs. */
export function mathTextState(node, reviewed = false) {
    if (node.namespaceURI !== MATH_NS) return { text: '', nativeText: null };
    const name = node.localName;
    if (name === 'r') {
        const text = runText(node);
        return { text, nativeText: reviewed ? text.replaceAll('-', '−').replaceAll("'", '′') : text };
    }
    const children = Array.from(node.children).filter((child) => !child.localName.endsWith('Pr'));
    const parts = children.map((child) => mathTextState(child, reviewed));
    const text = parts.map((part) => part.text).join('');
    const nativeText = parts.every((part) => part.nativeText !== null)
        ? parts.map((part) => part.nativeText).join('') : null;
    if (['oMath', 'oMathPara', 'e', 'sub', 'sup', 'num', 'den', 'lim'].includes(name)) return { text, nativeText };
    const argument = (key) => parts[children.findIndex((child) => child.localName === key)]?.text || '';
    const structured = parts.every((part) => part.nativeText !== null)
        ? (reviewed ? ',' : '\n') + parts.map((part) => part.nativeText).join(reviewed ? '-' : '\n') + (reviewed ? '.' : '\n') : null;
    if (['sSub', 'sSup', 'sSubSup'].includes(name)) {
        return { text: argument('e') + (name !== 'sSup' ? `_(${argument('sub')})` : '')
            + (name !== 'sSub' ? `^(${argument('sup')})` : ''), nativeText: structured };
    }
    if (name === 'f') {
        const type = attribute(property(property(node, 'fPr'), 'type'), 'val', 'bar');
        return { text: type === 'noBar' ? `binom(${argument('num')},${argument('den')})`
            : `(${argument('num')})/(${argument('den')})`, nativeText: reviewed ? null : type === 'bar' ? structured
                : type === 'noBar' && structured !== null ? '\n' + structured : null };
    }
    if (name === 'd') {
        const props = property(node, 'dPr');
        return { text: attribute(property(props, 'begChr'), 'val', '(')
            + parts.map((part) => part.text).join(attribute(property(props, 'sepChr'), 'val', '|'))
            + attribute(property(props, 'endChr'), 'val', ')'), nativeText: reviewed ? null : structured };
    }
    if (name === 'nary') return { text: attribute(property(property(node, 'naryPr'), 'chr'), 'val', '∫')
        + `_(${argument('sub')})^(${argument('sup')})(${argument('e')})`,
        nativeText: reviewed ? structured : nativeText === null ? null : '\n' + parts.slice(0, 2).map((part) => part.nativeText).join('\n')
            + (parts[2]?.nativeText || '') + '\n' };
    if (name === 'limLow' || name === 'limUpp') return { text: argument('e')
        + (name === 'limLow' ? '_' : '^') + `(${argument('lim')})`, nativeText: reviewed ? null : structured };
    // Keep unknown content readable, but do not authorize writes by assuming
    // its argument order or its Word text representation.
    return { text: text || '\uFFFC', nativeText: null };
}
