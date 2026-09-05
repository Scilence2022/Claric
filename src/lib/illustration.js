
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
 * Cap on the stored-SVG source embedded into a redesign prompt. Stored
 * sources were already sanitizeSvg-bounded at design time; this is a
 * defensive slice so a hand-inflated source cannot blow the context.
 */
const MAX_SOURCE_SVG_CHARS = 50 * 1024;

/**
 * Builds the LLM prompt that REDESIGNS an existing figure as one SVG
 * document (replace_illustration). Unlike buildIllustrationPrompt this
 * expects text: diagrams live on their labels and legends, so the
 * "avoid lettering" rule would defeat the task. When the original picture
 * is attached to the request, the prompt anchors fidelity to it — the
 * model must change only what the instruction asks for. When the stored
 * SVG source is available it is embedded verbatim and the task shifts from
 * "redraw" to "edit this markup minimally" — the lossless path.
 *
 * @param {string} instruction - What to change (not a full re-description)
 * @param {string} scopeText - Document text (subject/mood context)
 * @param {object} [options]
 * @param {boolean} [options.hasSourceImage=false] - True when the original
 *   picture rides along as an image input
 * @param {string} [options.sourceSvg=''] - Stored SVG source of the figure
 *   (svg-source-store), when the document carries it
 * @returns {string}
 */
export function buildIllustrationRedesignPrompt(instruction, scopeText, { hasSourceImage = false, sourceSvg = '' } = {}) {
    const svgSource = (typeof sourceSvg === 'string' && sourceSvg.trim().startsWith('<svg'))
        ? sourceSvg.trim().slice(0, MAX_SOURCE_SVG_CHARS)
        : '';
    let fidelity;
    if (svgSource && hasSourceImage) {
        fidelity = 'The CURRENT version of the figure is attached as an image, and its exact SVG ' +
            'source is included below. Edit that source minimally — reproduce its structure, flow, ' +
            'and every text label faithfully; change only what the instruction asks for.';
    } else if (svgSource) {
        fidelity = 'The exact SVG source of the figure to redesign is included below. Edit it ' +
            'minimally — keep its structure, flow, and every text label; change only what the ' +
            'instruction asks for.';
    } else if (hasSourceImage) {
        fidelity = 'The CURRENT version of the figure is attached as an image. Reproduce its structure, ' +
            'flow, and every text label faithfully — change only what the instruction asks for.';
    } else {
        fidelity = 'The figure to redesign is described in the user instruction below.';
    }
    return (
        'You are an illustrator embedded in Microsoft Word. Redesign an EXISTING figure of the ' +
        'document as a single self-contained SVG image.\n\n' +
        fidelity + '\n\n' +
        'OUTPUT CONTRACT (strict):\n' +
        '- Output ONLY the SVG markup: exactly one <svg ...>...</svg> element. No markdown, no code ' +
        'fences, no explanations, no commentary.\n' +
        '- The SVG must be self-contained: no external images, fonts, or URLs; no <script>, no event ' +
        'handlers, no <foreignObject>.\n' +
        '- Include width, height, and viewBox attributes on the root <svg> element.\n' +
        '- This is a diagram: text labels are expected. Render every label legibly — short strings, ' +
        'boxes sized to fit their text, no overlaps.\n' +
        '- Keep the original layout, color palette, and proportions unless the instruction says ' +
        'otherwise. Keep the markup under ~15 KB.\n' +
        '- Match the language of the document and of the original labels.\n\n' +
        'USER INSTRUCTION:\n' + (instruction || '').trim() + '\n\n' +
        (svgSource ? '--- CURRENT SVG SOURCE (edit this markup) ---\n' + svgSource + '\n\n' : '') +
        '--- DOCUMENT TEXT (context for subject and mood) ---\n' + (scopeText || '')
    );
}

/**
 * Label cap for extractSvgTextLabels results — enough for the edit tool's
 * error messages without dumping an entire pathological diagram.
 */
const MAX_SVG_LABELS = 40;
const MAX_SVG_LABEL_CHARS = 80;

/**
 * Parses an SVG and returns its text-bearing elements (text and tspan),
 * deepest first so nested tspans are edited before their parent text.
 *
 * @param {string} svg
 * @returns {{ doc: Document|null, nodes: Element[] }} doc is null when the
 *   markup cannot be parsed
 * @private
 */
function _parseSvgTextNodes(svg) {
    if (typeof svg !== 'string' || !svg.trim()) return { doc: null, nodes: [] };
    let doc;
    try {
        doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
    } catch {
        return { doc: null, nodes: [] };
    }
    if (doc.getElementsByTagName('parsererror').length > 0) return { doc: null, nodes: [] };
    const nodes = [];
    const walk = (el, depth) => {
        for (const child of Array.from(el.childNodes || [])) {
            if (child.nodeType !== 1) continue;
            if (child.localName === 'text' || child.localName === 'tspan') {
                walk(child, depth + 1);
                nodes.push({ el: child, depth });
            } else {
                walk(child, depth);
            }
        }
    };
    walk(doc.documentElement, 0);
    nodes.sort((a, b) => b.depth - a.depth);
    return { doc, nodes: nodes.map((n) => n.el) };
}

/**
 * Lists the visible text labels of an SVG (trimmed text/tspan contents,
 * document order, deduplicated). Used to tell the model which labels an
 * edit_illustration_text call could have matched.
 *
 * @param {string} svg
 * @returns {string[]}
 */
export function extractSvgTextLabels(svg) {
    const { nodes } = _parseSvgTextNodes(svg);
    const seen = new Set();
    const labels = [];
    // Shallowest last in node order; for label listing we want the parent
    // <text> aggregates, not their tspan fragments — walk deepest-first
    // nodes but keep only leaf text content per element's own text nodes.
    for (const el of nodes.slice().reverse()) {
        const own = Array.from(el.childNodes || [])
            .filter((n) => n.nodeType === 3)
            .map((n) => (/** @type {Text} */ (n)).data)
            .join('')
            .trim();
        if (!own) continue;
        const label = own.slice(0, MAX_SVG_LABEL_CHARS);
        if (seen.has(label)) continue;
        seen.add(label);
        labels.push(label);
        if (labels.length >= MAX_SVG_LABELS) break;
    }
    return labels;
}

/**
 * Applies deterministic find-and-replace edits to an SVG's text labels
 * (edit_illustration_text). Matching happens on DECODED text-node content
 * (entities resolved), so labels containing & or < match their drawn form;
 * only text nodes are touched — structure, styles, and geometry survive
 * verbatim. Each edit replaces every occurrence across the SVG.
 *
 * @param {string} svg - Stored SVG source
 * @param {Array<{old: string, new: string}>} edits - Exact label replacements
 * @returns {{ svg: string, applied: Array<{old: string, new: string, count: number}>,
 *   failed: Array<{old: string, new: string}>, labels: string[] }}
 *   svg is the edited markup (input unchanged when nothing applied);
 *   labels lists the labels present for retry guidance
 */
export function editSvgTextLabels(svg, edits) {
    const labels = extractSvgTextLabels(svg);
    const cleanEdits = (Array.isArray(edits) ? edits : [])
        .map((e) => ({
            old: e && e.old != null ? String(e.old) : '',
            new: e && e.new != null ? String(e.new) : '',
        }))
        .filter((e) => e.old.trim() && e.new.trim() && e.old !== e.new);
    if (cleanEdits.length === 0) return { svg, applied: [], failed: [], labels };

    const { doc, nodes } = _parseSvgTextNodes(svg);
    if (!doc) {
        return { svg, applied: [], failed: cleanEdits.map(({ old, new: n }) => ({ old, new: n })), labels };
    }

    const applied = [];
    const failed = [];
    for (const edit of cleanEdits) {
        let count = 0;
        for (const el of nodes) {
            for (const child of Array.from(el.childNodes || [])) {
                if (child.nodeType !== 3) continue;
                const textNode = /** @type {Text} */ (child);
                const hits = textNode.data.split(edit.old).length - 1;
                if (hits > 0) {
                    textNode.data = textNode.data.split(edit.old).join(edit.new);
                    count += hits;
                }
            }
        }
        if (count > 0) applied.push({ old: edit.old, new: edit.new, count });
        else failed.push({ old: edit.old, new: edit.new });
    }

    if (applied.length === 0) return { svg, applied, failed, labels };
    return { svg: new XMLSerializer().serializeToString(doc), applied, failed, labels };
}

/**
 * Max characters of document context fed to an image model. Image APIs cap
 * the prompt hard (OpenAI's images endpoint rejects very long prompts), and
 * a diagram brief needs the subject, not the whole document.
 */
const MAX_IMAGE_CONTEXT_CHARS = 1500;

/**
 * Builds the prompt for a TEXT-TO-IMAGE model (image-client.js), as opposed to
 * buildIllustrationPrompt which asks a chat LLM for SVG markup.
 *
 * The two differ in kind, not degree: an image model takes a description of
 * the picture and has no notion of an output contract, so the SVG rules
 * ("output only markup", "no foreignObject") would be noise at best and end up
 * rendered as literal text at worst. What matters here is a compact subject
 * description plus enough document context to match the topic, with an
 * explicit instruction to avoid lettering — every current image model renders
 * text unreliably, and a diagram full of garbled labels is worse than one with
 * none.
 *
 * @param {string} instruction - The user's request ("设计示意图并插入")
 * @param {string} scopeText - Document text for subject/topic context
 * @returns {string} A single prompt string for the image endpoint
 */
export function buildImagePrompt(instruction, scopeText) {
    const brief = (instruction || '').trim();
    const context = (scopeText || '').trim().slice(0, MAX_IMAGE_CONTEXT_CHARS);
    const parts = [
        'A clean, professional illustration suitable for inclusion in a formal document.',
        brief ? `Subject: ${brief}` : '',
        context ? `Document topic for context: ${context}` : '',
        'Style: clear, uncluttered, high contrast, neutral background, suitable for printing in a report.',
        'Do not render any words, letters, numbers, labels, or captions in the image.',
    ];
    return parts.filter(Boolean).join('\n\n');
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
 * Wording that pins the request to the SVG (chat-LLM) route regardless of
 * whether an image model is configured. A user asking for a vector graphic or
 * naming SVG explicitly wants markup — an image model cannot produce it.
 */
const SVG_REQUEST_RE = /\bsvg\b|矢量图?|向量图|line art|\bvector\b/i;

/**
 * Chooses which renderer designs an illustration.
 *
 * Two engines exist and they are good at different things: the chat LLM emits
 * SVG markup (crisp, tiny, but visibly synthetic for anything organic), while a
 * text-to-image model returns raster artwork (photographic or painterly, but
 * unreliable with text and heavier). The rule:
 *
 *   1. Explicit SVG/vector wording  -> 'svg' (an image model cannot deliver it)
 *   2. Image generation configured  -> 'image'
 *   3. Otherwise                    -> 'svg'
 *
 * Rule 2 makes configuring an image provider the switch that turns
 * "设计示意图并插入" into a real generated picture, with no extra syntax for
 * the user to remember, and rule 3 keeps every existing install working
 * exactly as before.
 *
 * @param {string} instruction - The user's illustration instruction
 * @param {boolean} imageModelReady - True when an image provider is enabled
 *   AND has an endpoint plus model configured
 * @returns {'svg' | 'image'}
 */
export function illustrationRenderer(instruction, imageModelReady) {
    if (SVG_REQUEST_RE.test(instruction || '')) return 'svg';
    return imageModelReady ? 'image' : 'svg';
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
