/** Compile formatting against observed native paragraphs before review. */
import { describeFormatParagraph, resolveFormatParagraphs } from '../lib/format-targets.js';
import { rangeStructureFingerprint } from '../lib/ooxml-fingerprint.js';
import { nativeHighlightColor, comparableHighlightColor } from '../lib/word-format-values.js';

const SCALARS = ['alignment', 'lineSpacing', 'spaceBefore', 'spaceAfter', 'leftIndent', 'rightIndent', 'firstLineIndent'];
const FONT_KEYS = ['bold', 'italic', 'underline', 'strikeThrough', 'doubleStrikeThrough', 'superscript',
    'subscript', 'color', 'highlightColor', 'name', 'size'];
const DESKTOP_FONT_KEYS = ['allCaps', 'smallCaps'];
const BATCH = 24;
const inside = (value) => ['Inside', 'InsideStart', 'InsideEnd', 'Equal'].includes(value);
const check = (signal) => { if (signal?.aborted) throw new DOMException('Formatting cancelled.', 'AbortError'); };
const key = (descriptor, mode) => mode === 'local-id' ? descriptor.nativeId : descriptor.id;
const canonical = (value) => typeof value === 'string' ? value.toLowerCase().replace(/[\s_-]+/g, '') : value;

function copyProperties(source, keys) {
    return Object.fromEntries(keys.map((name) => [name, source?.[name]]));
}

function desktopFontSupported() {
    try { return !!globalThis.Office?.context?.requirements?.isSetSupported?.('WordApiDesktop', '1.3'); }
    catch { return false; }
}

/** Per-paragraph XML avoids unreliable index mapping through nested objects. */
export async function readFormatInventory(context, scopeRange, { scope, signal, log = () => {} } = {}) {
    const collection = scopeRange.paragraphs;
    if (!collection || typeof collection.load !== 'function') return null;
    collection.load('items');
    await context.sync();
    check(signal);
    if (collection.items.length > 5000) throw new Error('Formatting exceeds the 5000-paragraph inspection limit.');
    const records = new Array(collection.items.length);
    async function readBatch(start, end) {
        const pending = [];
        try {
            check(signal);
            for (let index = start; index < end; index++) {
                const paragraph = collection.items[index];
                if (typeof paragraph.load !== 'function' || typeof paragraph.getOoxml !== 'function') {
                    throw new Error('Native paragraph evidence is unavailable.');
                }
                paragraph.load(['text', 'style', 'styleBuiltIn', ...SCALARS].join(','));
                const range = paragraph.getRange('Content');
                const table = paragraph.parentTableOrNullObject;
                if (!table || typeof table.load !== 'function') throw new Error('Native table membership is unavailable.');
                table.load('isNullObject');
                const relation = scope === 'document' ? null : range.compareLocationWith(scopeRange);
                const font = range.font;
                if (typeof font?.load === 'function') font.load(FONT_KEYS.join(','));
                pending.push({ index, paragraph, range, table, relation, font, xml: paragraph.getOoxml() });
            }
            await context.sync();
            check(signal);
            for (const item of pending) {
                const { index, paragraph, table, relation, font, xml } = item;
                const descriptor = describeFormatParagraph({ index, text: paragraph.text, style: paragraph.style,
                    styleBuiltIn: paragraph.styleBuiltIn, ooxml: xml.value,
                    inTable: typeof table.isNullObject === 'boolean' ? !table.isNullObject : undefined,
                    withinScope: scope === 'document' || inside(relation?.value), bold: font?.bold });
                records[index] = { descriptor, structure: rangeStructureFingerprint(xml.value),
                    paragraph: copyProperties(paragraph, SCALARS), font: copyProperties(font, FONT_KEYS) };
            }
        } catch (error) {
            check(signal);
            if (end - start > 1) {
                const mid = Math.floor((start + end) / 2);
                await readBatch(start, mid); await readBatch(mid, end);
            } else {
                let text;
                try { text = collection.items[start]?.text; }
                catch { /* A failed sync may leave even text unloaded. */ }
                const descriptor = describeFormatParagraph({ index: start, text, withinScope: false });
                records[start] = { descriptor, structure: null, paragraph: {}, font: {} };
                log(`Formatting paragraph ${start + 1} is read-only: ${error.message}`, 'warning');
            }
        }
    }
    for (let start = 0; start < collection.items.length; start += BATCH) {
        await readBatch(start, Math.min(start + BATCH, collection.items.length));
    }
    // Desktop-only font properties must never poison essential paragraph reads.
    // Keep failed optional reads isolated, even when capability detection says yes.
    if (desktopFontSupported()) {
        async function readDesktopFont(start, end) {
            try {
                const pending = [];
                for (let index = start; index < end; index++) {
                    if (!records[index].descriptor.verified) continue;
                    const font = collection.items[index].getRange('Content').font;
                    font.load(DESKTOP_FONT_KEYS.join(',')); pending.push({ index, font });
                }
                await context.sync(); check(signal);
                for (const { index, font } of pending) Object.assign(records[index].font, copyProperties(font, DESKTOP_FONT_KEYS));
            } catch (error) {
                check(signal);
                if (end - start > 1) {
                    const mid = Math.floor((start + end) / 2);
                    await readDesktopFont(start, mid); await readDesktopFont(mid, end);
                } else log(`Formatting paragraph ${start + 1}: optional font properties unavailable: ${error.message}`, 'info');
            }
        }
        for (let start = 0; start < collection.items.length; start += BATCH) {
            await readDesktopFont(start, Math.min(start + BATCH, collection.items.length));
        }
    }
    // Local IDs need Microsoft 365 in addition to WordApi 1.6. Keep this
    // optional capability read separate from the essential metadata batch.
    let identityMode = 'ordinal';
    if (typeof context.document.getParagraphByUniqueLocalId === 'function') {
        try {
            collection.items.forEach((paragraph) => paragraph.load('uniqueLocalId'));
            await context.sync(); check(signal);
            const ids = collection.items.map((paragraph) => paragraph.uniqueLocalId);
            if (ids.every((id) => typeof id === 'string' && id.length > 0) && new Set(ids).size === ids.length) {
                ids.forEach((id, index) => { records[index].descriptor.nativeId = id; });
                identityMode = 'local-id';
            }
        } catch (error) { check(signal); log(`Formatting uses verified paragraph order; local IDs unavailable: ${error.message}`, 'info'); }
    }
    return { identityMode, records, descriptors: records.map((record) => record.descriptor) };
}

function orderedIdentity(inventory) {
    return JSON.stringify(inventory.descriptors.map((d) => [d.text, d.styleBuiltIn, d.style, d.role, d.verified]));
}

function expectedEnum(group, property, value) {
    if (group === 'font' && property === 'highlightColor') return nativeHighlightColor(value);
    const enums = group === 'paragraph' && property === 'alignment' ? Word.Alignment
        : group === 'font' && property === 'underline' ? Word.UnderlineType
            : group === 'paragraph' && property === 'styleBuiltIn' ? (Word.BuiltInStyleName || Word.Style) : null;
    if (!enums) {
        if ((group === 'paragraph' && ['alignment', 'styleBuiltIn'].includes(property)) || property === 'underline') {
            throw new Error(`Word cannot verify ${property} formatting on this host.`);
        }
        return value;
    }
    const name = Object.keys(enums).find((entry) => canonical(entry) === canonical(value)
        || canonical(enums[entry]) === canonical(value));
    if (!name) throw new Error(`Unknown formatting ${property}: ${value}.`);
    return enums[name];
}

/** Structural/list and exact character-range changes keep the strict path. */
export function compileFormatTargetPlan(inventory, ops, { bodyOnly = false } = {}) {
    if (!inventory) {
        if (bodyOnly || ops.some((op) => op.paragraphRole || op.paragraphIds)) {
            throw new Error('Word could not inspect the requested formatting targets. No changes were applied.');
        }
        return null;
    }
    const actionable = ops.filter((op) => !op.insert && !op.cleanup);
    // Older hosts can still use exact scope/character formatting. Semantic
    // body selectors, however, must never fall back to that wider write.
    if (!bodyOnly && !actionable.some((op) => op.paragraphRole || op.paragraphIds)
        && !inventory.descriptors.some((d) => d.styleBuiltIn)) return null;
    // Reject unsupported values before staging, including the strict structural
    // path and inserted text. A warning at Apply must not silently drop an op.
    for (const op of ops) {
        for (const group of ['paragraph', 'font']) {
            for (const [property, value] of Object.entries(op[group] || {})) {
                expectedEnum(group, property, value);
                if (group === 'font' && DESKTOP_FONT_KEYS.includes(property) && !desktopFontSupported()) {
                    throw new Error(`Word cannot verify ${property} formatting on this host.`);
                }
            }
        }
    }
    const mode = ops.some((op) => op.insert || op.cleanup || op.match && op.font
        || Object.keys(op.paragraph || {}).some((property) => !SCALARS.includes(property))) ? 'strict' : 'targets';
    const entries = [];
    const baselines = {};
    const final = {};
    const excluded = new Map();
    for (const op of actionable) {
        if (op.match?.length > 255) throw new Error('Formatting match exceeds Word search limits. Use a paragraph ID instead.');
        const selected = resolveFormatParagraphs(inventory.descriptors, op, { bodyOnly });
        if (!selected.targets.length) throw new Error('No verified formatting paragraphs matched. Choose an explicit target or revise the request.');
        selected.exclusions.forEach((entry) => excluded.set(entry.id, entry));
        const ids = selected.targets.map((d) => key(d, inventory.identityMode));
        entries.push({ signature: JSON.stringify(op), ids });
        for (const d of selected.targets) {
            const id = key(d, inventory.identityMode);
            const record = inventory.records[d.index];
            baselines[id] = { descriptor: d, structure: record.structure, paragraph: record.paragraph, font: record.font };
            final[id] ||= { paragraph: {}, font: {} };
            for (const group of ['paragraph', 'font']) {
                for (const [property, value] of Object.entries(op[group] || {})) {
                    if ((group === 'font' || SCALARS.includes(property)) && record[group][property] === undefined) {
                        throw new Error(`Word cannot verify ${property} formatting on this host.`);
                    }
                    if (mode === 'targets') final[id][group][property] = expectedEnum(group, property, value);
                }
            }
        }
    }
    return { mode, bodyOnly, identityMode: inventory.identityMode, orderedIdentity: orderedIdentity(inventory),
        entries, baselines, final, summary: { verifiedParagraphs: Object.keys(baselines).length,
            excludedParagraphs: excluded.size, exclusions: [...excluded.values()], identityMode: inventory.identityMode } };
}

export function validateFormatTargetPlan(plan, inventory, ops) {
    if (!inventory || inventory.identityMode !== plan.identityMode) throw new Error('Formatting paragraph identity could not be verified. Draft a new proposal.');
    if (plan.identityMode === 'ordinal' && orderedIdentity(inventory) !== plan.orderedIdentity) {
        throw new Error('Formatting paragraph order or scope changed. Draft a new proposal.');
    }
    const current = compileFormatTargetPlan(inventory, ops, { bodyOnly: plan.bodyOnly });
    if (!current) throw new Error('Formatting target evidence is unavailable. Draft a new proposal.');
    for (const entry of current.entries) {
        const expected = plan.entries.find((prior) => prior.signature === entry.signature);
        if (!expected || JSON.stringify(expected.ids) !== JSON.stringify(entry.ids)) {
            throw new Error('The verified formatting target set changed. Draft a new proposal.');
        }
        for (const id of entry.ids) {
            const before = plan.baselines[id]; const now = current.baselines[id];
            if (!before || before.descriptor.text !== now.descriptor.text
                || before.descriptor.styleBuiltIn !== now.descriptor.styleBuiltIn
                || before.descriptor.style !== now.descriptor.style || before.descriptor.role !== now.descriptor.role
                || before.structure !== now.structure) {
                throw new Error('A verified formatting paragraph changed. Draft a new proposal.');
            }
            for (const group of ['paragraph', 'font']) {
                for (const property of Object.keys(current.final[id][group])) {
                    if (!propertyEqual(before[group][property], now[group][property], property)) {
                        throw new Error(`Formatting ${property} baseline changed. Draft a new proposal.`);
                    }
                }
            }
        }
    }
    return current;
}

function propertyEqual(actual, expected, property) {
    if (typeof actual === 'number' && typeof expected === 'number') return Math.abs(actual - expected) <= 0.051;
    if (property === 'highlightColor') return comparableHighlightColor(actual) === comparableHighlightColor(expected);
    const normalize = (value) => typeof value !== 'string' ? value
        : ['alignment', 'underline'].includes(property) ? canonical(value) : value.toLowerCase();
    return normalize(actual) === normalize(expected);
}

/** Absolute writes, one final expectation per property, followed by native read-back. */
export async function applyVerifiedFormatTargets(context, scopeRange, plan, inventory, anchor, { signal } = {}) {
    const collection = scopeRange.paragraphs;
    const pending = [];
    let noopParagraphs = 0;
    for (const [id, expected] of Object.entries(plan.final)) {
        const descriptor = inventory.descriptors.find((d) => key(d, plan.identityMode) === id);
        if (!descriptor) throw new Error('A verified formatting target disappeared.');
        const paragraph = collection.items[descriptor.index];
        const font = paragraph.getRange('Content').font;
        const before = inventory.records[descriptor.index];
        const changes = [];
        for (const group of ['paragraph', 'font']) {
            for (const [property, value] of Object.entries(expected[group])) {
                if (!propertyEqual(before[group][property], value, property)) changes.push({ group, property, value });
            }
        }
        if (!changes.length) noopParagraphs++;
        pending.push({ paragraph, font, descriptor, expected, changes });
    }
    check(signal);
    for (const item of pending) {
        for (const { group, property, value } of item.changes) {
            check(signal);
            anchor.attempted = true;
            (group === 'paragraph' ? item.paragraph : item.font)[property] = value;
        }
    }
    await context.sync(); check(signal);
    for (const item of pending) {
        item.paragraph.load(['text', ...Object.keys(item.expected.paragraph)].join(','));
        if (Object.keys(item.expected.font).length) item.font.load(Object.keys(item.expected.font).join(','));
    }
    await context.sync(); check(signal);
    for (const item of pending) {
        if (item.paragraph.text !== item.descriptor.text) throw new Error('Formatting read-back detected changed paragraph text. Review the document.');
        for (const group of ['paragraph', 'font']) {
            for (const [property, value] of Object.entries(item.expected[group])) {
                const actual = (group === 'paragraph' ? item.paragraph : item.font)[property];
                if (!propertyEqual(actual, value, property)) throw new Error(`Word read-back did not confirm ${property}. Review the document.`);
            }
        }
    }
    anchor.attempted = true; // consume verified no-ops too; never replay a stale card
    return { applied: pending.length > noopParagraphs, appliedRanges: pending.length - noopParagraphs,
        verifiedParagraphs: pending.length, noopParagraphs, alreadySatisfied: pending.length === noopParagraphs };
}
