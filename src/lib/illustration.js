
import { getPurifier } from './sanitize.js';

/**
 * Illustration Module
 *
 * Illustration turns ("设计并增加SVG插图", "给文章配一张插图") ask the chat
 * LLM to design artwork as a single self-contained SVG, which is sanitized
 * here and rasterized to PNG by the caller (word-actions) for insertion —
 * Word's insertInlinePictureFromBase64 does not accept SVG, but takes PNG
 * (and JPEG/GIF/BMP) base64, which covers the mainstream raster formats.
 *
 * Pure module except sanitizeSvg (lazy DOMPurify, DOM needed at call time).
 * No Word API usage. Safe to import under Jest/node.
 *
 * @module illustration
 */

/**
 * SVG payloads beyond this size are pathological for a chat LLM and almost
 * always truncated mid-document — reject rather than insert broken XML.
 */
const MAX_SVG_CHARS = 256 * 1024;

/**
 * Builds the LLM prompt that designs an illustration as one SVG document.
 *
 * @param {string} instruction - The user's illustration instruction
 * @param {string} scopeText - Document text (subject/mood context)
 * @returns {string}
 */
export function buildIllustrationPrompt(instruction, scopeText) {
    return (
        'You are an illustrator embedded in Microsoft Word. Design ONE illustration for the document ' +
        'below and output it as a single self-contained SVG image.\n\n' +
        'OUTPUT CONTRACT (strict):\n' +
        '- Output ONLY the SVG markup: exactly one <svg ...>...</svg> element. No markdown, no code ' +
        'fences, no explanations, no commentary.\n' +
        '- The SVG must be self-contained: no external images, fonts, or URLs; no <script>, no event ' +
        'handlers, no <foreignObject>.\n' +
        '- Include width, height, and viewBox attributes on the root <svg> element.\n' +
        '- Aim for a polished, coherent scene: layered shapes, gradients, and subtle texture rather ' +
        'than a few primitive boxes. Keep the markup under ~15 KB.\n' +
        '- Avoid embedding text or letters in the artwork — they render unreliably.\n' +
        '- Match the subject matter, mood, and language of the document.\n\n' +
        'USER INSTRUCTION:\n' + (instruction || '').trim() + '\n\n' +
        '--- DOCUMENT TEXT (context for subject and mood) ---\n' + (scopeText || '')
    );
}

/**
 * Extracts the SVG document from the model's raw output. Tolerates code
 * fences and surrounding prose by slicing from the first "<svg" to the
 * first "</svg>". Returns null (with a warning) when nothing usable is
 * found or the payload is pathologically large.
 *
 * @param {string} raw - Raw model output
 * @param {function} [log] - Logging callback
 * @returns {{ svg: string } | null}
 */
export function parseIllustration(raw, log = () => {}) {
    if (!raw) return null;
    let text = String(raw).trim();

    const fence = text.match(/```(?:svg|xml|html)?\s*([\s\S]*?)```/);
    if (fence) text = fence[1].trim();

    const start = text.indexOf('<svg');
    const end = text.indexOf('</svg>', start === -1 ? 0 : start);
    if (start === -1 || end === -1) {
        log('Illustration: no <svg> element found in the model response', 'warning');
        return null;
    }
    const svg = text.slice(start, end + '</svg>'.length);
    if (svg.length > MAX_SVG_CHARS) {
        log(`Illustration: SVG payload too large (${svg.length} chars, cap ${MAX_SVG_CHARS})`, 'warning');
        return null;
    }
    return { svg };
}

/**
 * Sanitizes an SVG string for safe DOM rendering (preview) and insertion.
 * Keeps the SVG vocabulary (shapes, gradients, filters) while stripping
 * scripts, event handlers, foreignObject, iframes, and other active
 * content — the markup is raw LLM output and may echo prompt-injected
 * document content.
 *
 * @param {string} svg - SVG markup from parseIllustration
 * @returns {string} Sanitized SVG markup
 */
export function sanitizeSvg(svg) {
    return getPurifier().sanitize(svg, {
        USE_PROFILES: { svg: true, svgFilters: true },
        FORBID_TAGS: ['foreignObject', 'iframe', 'script'],
    });
}

/**
 * Reads the declared pixel dimensions of an SVG: explicit width/height
 * attributes first (units ignored), then the viewBox extent. Returns null
 * when neither is present.
 *
 * @param {string} svg - SVG markup
 * @returns {{ width: number, height: number } | null}
 */
export function svgDimensions(svg) {
    const root = String(svg || '').match(/<svg\b[^>]*>/i);
    if (!root) return null;
    const tag = root[0];
    const attr = (name) => {
        const m = tag.match(new RegExp(`${name}\\s*=\\s*"([^"]+)"`, 'i'))
            || tag.match(new RegExp(`${name}\\s*=\\s*'([^']+)'`, 'i'));
        return m ? parseFloat(m[1]) : NaN;
    };
    const width = attr('width');
    const height = attr('height');
    if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
        return { width, height };
    }
    const vb = tag.match(/viewBox\s*=\s*["']([^"']+)["']/i);
    if (vb) {
        const parts = vb[1].trim().split(/[\s,]+/).map(Number);
        if (parts.length === 4 && parts.every(Number.isFinite) && parts[2] > 0 && parts[3] > 0) {
            return { width: parts[2], height: parts[3] };
        }
    }
    return null;
}

/**
 * Ensures the root <svg> carries width/height attributes — browsers
 * rasterize dimension-less SVGs at a fallback size (often 300×150), which
 * would blur the inserted image. Values come from svgDimensions (viewBox
 * included); missing everything yields a 1200×800 default.
 *
 * @param {string} svg - SVG markup
 * @returns {string} SVG markup with width/height on the root element
 */
export function ensureSvgDimensions(svg) {
    const text = String(svg || '');
    if (/<svg\b[^>]*\swidth\s*=/i.test(text) && /<svg\b[^>]*\sheight\s*=/i.test(text)) return text;
    const dims = svgDimensions(text) || { width: 1200, height: 800 };
    return text.replace(/<svg\b/i, `<svg width="${dims.width}" height="${dims.height}"`);
}

/**
 * Picks the insertion position from the instruction's own wording:
 * cursor-anchored phrasings (光标/此处/当前位置...) insert at the caret;
 * hero-image phrasings (题图/头图/开头...) land at the document start;
 * everything else appends at the end. Cursor wins over start/end when
 * both match — it is the explicit anchor ("在光标处插一张题图" means the
 * caret, not the document top).
 *
 * The position resolves to an anchor at APPLY time: 'cursor' reads the
 * selection when the user clicks Apply (Word keeps the document selection
 * while focus is in the taskpane), so the image lands where the caret
 * actually is, not where it was when the request was typed.
 *
 * @param {string} text - Trimmed chat input
 * @returns {'start' | 'end' | 'cursor'}
 */
export function illustrationPositionFromInstruction(text) {
    const CURSOR_RE = /光标|当前(?:插入)?位置|此处|这里|\b(?:cursor|caret)\b|\bcurrent position\b|\binsertion point\b|\bhere\b/i;
    if (CURSOR_RE.test(text || '')) return 'cursor';
    const START_RE = /开头|顶部|文首|卷首|扉页|题图|头图|标题下|\btop\b|\bheader\b|\bbeginning\b/i;
    return START_RE.test(text || '') ? 'start' : 'end';
}

/**
 * Human-readable label for an illustration position, for proposal-card
 * strings ("SVG 8.5 KB → PNG at the cursor").
 *
 * @param {'start' | 'end' | 'cursor'} position
 * @returns {string}
 */
export function illustrationPositionLabel(position) {
    if (position === 'start') return 'document start';
    if (position === 'cursor') return 'the cursor';
    return 'document end';
}
