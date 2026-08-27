/**
 * Image Tool Model
 *
 * L2 of the tool-calling stack: the draft model image tools operate on.
 * Seeded from a snapshot of the document's inline pictures. Content ops
 * (design/replace) execute a nested LLM call host-side; visual READING is
 * the host-executed read_image tool (agent-actions attaches the picture to
 * the next loop message as a multimodal image input — backend-dependent,
 * degraded gracefully by the send wrapper). Ops accumulate as a transaction
 * applied by applyImageOps only when the user clicks Apply on the proposal
 * card.
 *
 * Index discipline: image indexes are STABLE snapshot indexes (1-based,
 * document order at prepare time). Deletes/replaces mark their slot
 * consumed; inserts never shift snapshot indexes. Apply resolves indexes
 * against a re-read picture list with a staleness guard on the count.
 *
 * Pure module — no Office.js, no network. Hermetic-testable.
 *
 * @module image-model
 */

import { defineTool } from './tool-registry.js';

/** Max inline pictures listed into the loop (document-scale guard). */
const MAX_LISTED_IMAGES = 50;

/** Max width for the resize tool, in points — matches the insert cap. */
const MAX_IMAGE_WIDTH_PT = 450;

/** Positions an inserted image can land at. */
export const IMAGE_POSITIONS = Object.freeze(['start', 'end', 'cursor']);

/**
 * Tool specs for the image draft model. design/replace execute a nested LLM
 * call host-side (the illustration design prompt) — the loop model only
 * orchestrates.
 */
export const IMAGE_TOOL_SPECS = Object.freeze([
    defineTool({
        name: 'list_images',
        description: 'List the document\'s inline pictures with their 1-based index, size in points, and alt text. Indexes refer to this snapshot for every other image tool.',
        argsExample: {},
    }),
    defineTool({
        name: 'read_image',
        description: 'View the VISUAL CONTENT of one image. The picture is attached as an image input to the observation that follows this call, so vision-capable models see it directly. Requires a vision-capable model — if the observation reports the backend rejected image input, fall back to alt text and metadata and say so. "index" is the snapshot index of the picture.',
        argsExample: { index: 1 },
    }),
    defineTool({
        name: 'design_illustration',
        description: 'Design ONE new illustration for the document and stage it for insertion. "instruction" is the design brief (subject, style, mood); the host generates the SVG. "position" is where it inserts: "start" (document top), "end" (document bottom), or "cursor" (at the caret when Apply is clicked).',
        argsExample: { instruction: 'a minimalist line drawing of a rising sun', position: 'end' },
    }),
    defineTool({
        name: 'replace_illustration',
        description: 'Replace one existing picture with a newly designed illustration. "index" is the snapshot index of the picture to replace; "instruction" is the design brief for its replacement.',
        argsExample: { index: 2, instruction: 'the same scene but at dusk' },
    }),
    defineTool({
        name: 'delete_image',
        description: 'Delete one existing picture by its snapshot index.',
        argsExample: { index: 1 },
    }),
    defineTool({
        name: 'resize_image',
        description: 'Scale one existing picture to a new width in points (height follows the aspect ratio). Width is capped at 450pt.',
        argsExample: { index: 1, widthPt: 300 },
    }),
    defineTool({
        name: 'set_alt_text',
        description: 'Set the alt-text description of one existing picture (accessibility).',
        argsExample: { index: 1, text: 'Quarterly revenue chart' },
    }),
]);

/**
 * Creates the image draft model.
 *
 * @param {Array<{index: number, width: number, height: number, altText: string}>} snapshot -
 *   Inline-picture metadata in document order (1-based index)
 * @returns {{listImages: Function, recordInsert: Function, recordReplace: Function,
 *   recordDelete: Function, recordResize: Function, recordAltText: Function,
 *   describeOps: Function, ops: Array<object>, imageCount: number}}
 */
export function createImageModel(snapshot) {
    const images = (Array.isArray(snapshot) ? snapshot : []).slice(0, MAX_LISTED_IMAGES);
    const byIndex = new Map(images.map((img) => [img.index, img]));
    const consumed = new Set();
    /** @type {Array<object>} */
    const ops = [];

    const _err = (error) => ({ ok: false, error });
    const _ok = (result) => ({ ok: true, result });

    /** Indexes that still exist in the draft (not deleted/replaced). */
    const live = (index) => byIndex.has(index) && !consumed.has(index);

    function describe(index) {
        /** @type {{width?: number, height?: number, altText?: string}} */
        const img = byIndex.get(index) || {};
        return `image ${index} (${img.width}x${img.height}pt${img.altText ? `, alt "${String(img.altText).slice(0, 40)}"` : ''})`;
    }

    function listImages() {
        return _ok({
            count: images.length,
            images: images.map((img) => ({
                index: img.index,
                widthPt: img.width,
                heightPt: img.height,
                altText: img.altText || '',
            })),
            pendingOps: ops.map((op, i) => ({ id: i + 1, ...op, svg: undefined })),
        });
    }

    function recordInsert({ position, instruction, svg }) {
        if (!IMAGE_POSITIONS.includes(position)) {
            return _err(`"position" must be one of ${IMAGE_POSITIONS.join('|')}.`);
        }
        if (typeof instruction !== 'string' || !instruction.trim()) {
            return _err('"instruction" must be a non-empty design brief string.');
        }
        if (typeof svg !== 'string' || !svg.trim().startsWith('<svg')) {
            return _err('The design step produced no usable SVG.');
        }
        ops.push({ type: 'insert', position, instruction: instruction.trim(), svg });
        return _ok({ staged: `insert at ${position}`, svgKb: (svg.length / 1024).toFixed(1) });
    }

    function recordReplace({ index, instruction, svg }) {
        if (!Number.isInteger(index) || !live(index)) {
            return _err(`"index" must be a live snapshot index (1..${images.length}, not already deleted or replaced).`);
        }
        if (typeof instruction !== 'string' || !instruction.trim()) {
            return _err('"instruction" must be a non-empty design brief string.');
        }
        if (typeof svg !== 'string' || !svg.trim().startsWith('<svg')) {
            return _err('The design step produced no usable SVG.');
        }
        consumed.add(index);
        ops.push({ type: 'replace', index, instruction: instruction.trim(), svg });
        return _ok({ staged: `replace ${describe(index)}` });
    }

    function recordDelete(index) {
        if (!Number.isInteger(index) || !live(index)) {
            return _err(`"index" must be a live snapshot index (1..${images.length}, not already deleted or replaced).`);
        }
        consumed.add(index);
        ops.push({ type: 'delete', index });
        return _ok({ staged: `delete ${describe(index)}` });
    }

    function recordResize(index, widthPt) {
        if (!Number.isInteger(index) || !live(index)) {
            return _err(`"index" must be a live snapshot index (1..${images.length}, not already deleted or replaced).`);
        }
        if (!Number.isFinite(widthPt) || widthPt < 8 || widthPt > MAX_IMAGE_WIDTH_PT) {
            return _err(`"widthPt" must be a number between 8 and ${MAX_IMAGE_WIDTH_PT}.`);
        }
        ops.push({ type: 'resize', index, widthPt: Math.round(widthPt) });
        return _ok({ staged: `resize ${describe(index)} to ${Math.round(widthPt)}pt wide` });
    }

    function recordAltText(index, text) {
        if (!Number.isInteger(index) || !live(index)) {
            return _err(`"index" must be a live snapshot index (1..${images.length}, not already deleted or replaced).`);
        }
        if (typeof text !== 'string' || !text.trim()) {
            return _err('"text" must be a non-empty string.');
        }
        ops.push({ type: 'altText', index, text: text.trim().slice(0, 200) });
        return _ok({ staged: `alt text for ${describe(index)}` });
    }

    /**
     * Proposal-card change items for the recorded ops.
     *
     * @returns {Array<{id: number, label: string, before: string, after: string, svg?: string}>}
     */
    function describeOps() {
        return ops.map((op, i) => {
            const id = i + 1;
            const base = { id };
            switch (op.type) {
                case 'insert':
                    return { ...base, label: `Insert illustration at ${op.position}`, before: '', after: `${op.instruction} (${(op.svg.length / 1024).toFixed(1)} KB SVG → PNG)`, svg: op.svg };
                case 'replace':
                    return { ...base, label: `Replace ${describe(op.index)}`, before: byIndex.get(op.index)?.altText || 'existing picture', after: `${op.instruction} (${(op.svg.length / 1024).toFixed(1)} KB SVG → PNG)`, svg: op.svg };
                case 'delete':
                    return { ...base, label: `Delete ${describe(op.index)}`, before: 'existing picture', after: '' };
                case 'resize':
                    return { ...base, label: `Resize ${describe(op.index)}`, before: `${byIndex.get(op.index)?.width}pt wide`, after: `${op.widthPt}pt wide` };
                case 'altText':
                    return { ...base, label: `Alt text for ${describe(op.index)}`, before: byIndex.get(op.index)?.altText || '', after: op.text };
                default:
                    return { ...base, label: `Unknown op ${op.type}`, before: '', after: '' };
            }
        });
    }

    return {
        listImages,
        recordInsert,
        recordReplace,
        recordDelete,
        recordResize,
        recordAltText,
        describeOps,
        ops,
        get imageCount() { return images.length; },
    };
}
