import { PAGE_FIELDS, LINE_FIELDS, validatePageLayoutPatch, parsePageLayoutOps, explicitLineNumberRemoval, buildPageLayoutPrompt, describeLayoutChange } from '../lib/page-layout.js';
import { getActiveBackendConfig } from './app-state.js';
import { documentPartRoot } from '../lib/ooxml-text.js';
import { rangeStructureFingerprint } from '../lib/ooxml-fingerprint.js';

const proposals = new WeakMap();
const check = (signal) => { if (signal?.aborted) throw new DOMException('Page layout cancelled.', 'AbortError'); };
const signature = (op) => JSON.stringify({ sectionId: op.sectionId, patch: validatePageLayoutPatch(op.patch) });

export function requirePageLayout() {
    if (!globalThis.Office?.context?.requirements?.isSetSupported?.('WordApiDesktop', '1.3')) {
        throw new Error('Page layout requires WordApiDesktop 1.3 on Word for Mac or Windows. This host does not expose the required native APIs; no changes were applied.');
    }
}

/** PageSetup changes legitimately alter sectPr, but never the body content. */
function bodyFingerprint(xml) {
    const root = documentPartRoot(xml);
    if (!root) throw new Error('Cannot verify section body structure.');
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    for (const name of ['sectPr', 'sectPrChange']) {
        for (const element of Array.from(root.getElementsByTagNameNS(ns, name))) element.remove();
    }
    const result = rangeStructureFingerprint(new XMLSerializer().serializeToString(root));
    if (!result) throw new Error('Cannot fingerprint the section body.');
    return result;
}

/**
 * Native reads, one per property. Word for Mac rejects some PageSetup reads
 * (e.g. 0x80020004) and reports others as missing; a rejected read is probed
 * once, remembered in `unreadable`, and left out instead of failing the task.
 */
const READS = [
    { key: 'section.protectedForForms', queue: (e) => e.section.load('protectedForForms') },
    ...PAGE_FIELDS.map((name) => ({ key: `pageSetup.${name}`, queue: (e) => e.setup.load(name) })),
    ...LINE_FIELDS.map((name) => ({ key: `lineNumbering.${name}`, queue: (e) => e.setup.lineNumbering.load(name) })),
    { key: 'columns.count', queue: (e) => e.columns.load('items') },
    { key: 'columns.width', queue: (e) => e.columns.load('items/width') },
    { key: 'columns.spaceAfter', queue: (e) => e.columns.load('items/spaceAfter') },
    { key: 'columns.evenlySpaced', queue: (e) => { e.flags.evenlySpaced = e.columns.getIsEvenlySpaced(); } },
    { key: 'columns.lineBetween', queue: (e) => { e.flags.lineBetween = e.columns.getHasLineBetween(); } },
    { key: 'columns.flowDirection', queue: (e) => { e.flags.flowDirection = e.columns.getFlowDirection(); } },
];
const scalar = (value) => ['string', 'boolean'].includes(typeof value) || typeof value === 'number' && Number.isFinite(value);

function located(error) {
    const where = error?.debugInfo?.errorLocation;
    if (where && error instanceof Error && !error.message.includes(where)) error.message += ` (at ${where})`;
    return error;
}

async function sync(context, signal) {
    try { await context.sync(); } catch (error) { throw located(error); }
    check(signal);
}

async function loadReadable(context, entries, unreadable, signal) {
    const queue = (reads) => { for (const read of reads) for (const entry of entries) read.queue(entry); };
    try {
        queue(READS.filter((read) => !unreadable.has(read.key)));
        await context.sync(); check(signal);
    } catch (error) {
        if (error?.name === 'AbortError') throw error;
        let found = false;
        for (const read of READS.filter((r) => !unreadable.has(r.key))) {
            queue([read]);
            try { await context.sync(); } catch { unreadable.add(read.key); found = true; }
            check(signal);
        }
        if (!found) throw located(error);
    }
}

function readable(object, group, names, unreadable) {
    return Object.fromEntries(names.filter((name) => !unreadable.has(`${group}.${name}`) && scalar(object[name])).map((name) => [name, object[name]]));
}

function record(entry, unreadable) {
    const columns = {};
    if (!unreadable.has('columns.count')) {
        columns.count = entry.columns.items.length;
        columns.items = entry.columns.items.map((column) => readable(column, 'columns', ['width', 'spaceAfter'], unreadable));
    }
    for (const [name, result] of Object.entries(entry.flags)) {
        const value = unreadable.has(`columns.${name}`) ? undefined : result.value;
        if (name === 'flowDirection' ? ['LeftToRight', 'RightToLeft'].includes(value) : typeof value === 'boolean') columns[name] = value;
    }
    const protectedForForms = unreadable.has('section.protectedForForms') ? undefined : entry.section.protectedForForms;
    return { id: entry.id, ...(typeof protectedForForms === 'boolean' ? { protectedForForms } : {}),
        bodyFingerprint: bodyFingerprint(entry.xml.value), pageSetup: readable(entry.setup, 'pageSetup', PAGE_FIELDS, unreadable),
        lineNumbering: readable(entry.setup.lineNumbering, 'lineNumbering', LINE_FIELDS, unreadable), columns,
        selected: entry.relations ? entry.relations.some((r) => r.value === 'Equal') : true };
}

async function snapshot(context, { selection = false, signal, unreadable = new Set() } = {}) {
    check(signal);
    const sections = context.document.sections;
    sections.load('items');
    const selected = selection ? context.document.getSelection().sections : null;
    if (selected) selected.load('items');
    await sync(context, signal);
    if (!sections.items.length || sections.items.length > 100) throw new Error('Page layout supports 1–100 native sections per request.');
    const records = [];
    const handles = [];
    for (let start = 0; start < sections.items.length; start += 12) {
        const batch = sections.items.slice(start, start + 12).map((section, offset) => {
            const setup = section.pageSetup;
            if (!setup?.load || !setup.lineNumbering?.load || !setup.textColumns?.load) throw new Error('Native section page setup is unavailable on this host.');
            return { section, setup, columns: setup.textColumns, flags: {}, id: `s${start + offset + 1}` };
        });
        await loadReadable(context, batch, unreadable, signal);
        for (const entry of batch) {
            entry.relations = selected ? selected.items.map((target) => entry.section.body.getRange('Start').compareLocationWith(target.body.getRange('Start'))) : null;
            entry.xml = entry.section.body.getOoxml();
        }
        await sync(context, signal);
        for (const entry of batch) {
            handles.push(entry.section);
            records.push(record(entry, unreadable));
        }
    }
    const allowed = records.filter((r) => r.selected);
    if (selected && (allowed.length !== selected.items.length || !allowed.length)) throw new Error('Cannot identify the selected native sections safely.');
    return { records, handles, allowed, unreadable };
}

/** Compares only what every snapshot so far could read, so a newly unreadable property is not a stale document. */
function project({ selected: _selected, ...r }, unreadable) {
    const drop = (group, object) => Object.fromEntries(Object.entries(object).filter(([key]) => !unreadable.has(`${group}.${key}`)));
    const { count, items, ...flags } = r.columns;
    return { ...r, protectedForForms: unreadable.has('section.protectedForForms') ? undefined : r.protectedForForms,
        pageSetup: drop('pageSetup', r.pageSetup), lineNumbering: drop('lineNumbering', r.lineNumbering),
        columns: { ...drop('columns', flags), ...(unreadable.has('columns.count') ? {} : { count, items: items?.map((item) => drop('columns', item)) }) } };
}
const snapshotKey = (records, unreadable) => JSON.stringify(records.map((r) => project(r, unreadable)));
const publicRecord = ({ bodyFingerprint: _body, selected: _selected, ...r }) => r;

/** A setting can be changed only when Word can read it back for verification. */
function requireReadable(before, patch) {
    for (const [group, properties] of Object.entries(patch)) for (const [key, value] of Object.entries(properties)) {
        const known = group === 'columns' && key === 'items'
            ? before.columns.count !== undefined && value.every((item) => Object.keys(item).every((name) => before.columns.items.every((c) => c[name] !== undefined)))
            : before[group][key] !== undefined;
        if (!known) throw new Error(`Word on this host cannot read ${group === 'lineNumbering' ? 'line numbering' : group} ${key} for section ${before.id.slice(1)}, so the change cannot be verified. No layout changes were prepared.`);
    }
}

function validateGeometry(before, patch) {
    if (!patch.pageSetup && !patch.columns) return;
    const page = { ...before.pageSetup, ...patch.pageSetup };
    const needed = ['pageWidth', 'pageHeight', 'topMargin', 'bottomMargin', 'leftMargin', 'rightMargin', 'gutter',
        ...(page.gutter > 0 ? ['gutterPosition'] : []), ...(patch.pageSetup?.orientation || patch.pageSetup?.paperSize ? ['orientation'] : [])];
    const missing = needed.filter((name) => page[name] === undefined);
    if (missing.length) throw new Error(`Word could not read ${missing.join(', ')} for section ${before.id.slice(1)}; page geometry cannot be checked safely.`);
    const sizes = { A3: [841.89, 1190.55], A4: [595.28, 841.89], A5: [419.53, 595.28], Letter: [612, 792], Legal: [612, 1008], Tabloid: [792, 1224], Ledger: [1224, 792] };
    let width = page.pageWidth, height = page.pageHeight;
    if (patch.pageSetup?.paperSize && sizes[page.paperSize]) [width, height] = sizes[page.paperSize];
    if (patch.pageSetup?.orientation || patch.pageSetup?.paperSize) {
        [width, height] = page.orientation === 'Landscape' ? [Math.max(width, height), Math.min(width, height)] : [Math.min(width, height), Math.max(width, height)];
    }
    const gutterX = page.gutterPosition === 'Top' ? 0 : page.gutter;
    const gutterY = page.gutterPosition === 'Top' ? page.gutter : 0;
    const contentWidth = width - page.leftMargin - page.rightMargin - gutterX;
    if (contentWidth <= 0 || height <= page.topMargin + page.bottomMargin + gutterY) throw new Error('Requested margins and gutter leave no printable body area.');
    if (patch.columns?.items) {
        const total = patch.columns.items.reduce((sum, item, i, all) => sum + (item.width ?? before.columns.items[i]?.width ?? contentWidth / all.length)
            + (i < all.length - 1 ? item.spaceAfter ?? before.columns.items[i]?.spaceAfter ?? 0 : 0), 0);
        if (total > contentWidth + 1) throw new Error('Requested column widths and gaps exceed the body width.');
        if (patch.columns.items.length !== (patch.columns.count ?? before.columns.count)) throw new Error('Column dimensions must cover every targeted column.');
    }
}

function satisfied(value, expected) {
    if (typeof expected === 'number') return typeof value === 'number' && Math.abs(value - expected) < 0.1;
    if (Array.isArray(expected)) return Array.isArray(value) && value.length === expected.length && expected.every((item, i) => satisfied(value[i], item));
    if (expected && typeof expected === 'object') return Object.entries(expected).every(([key, item]) => satisfied(value?.[key], item));
    return value === expected;
}

export async function readPageLayout(_deps, { scope = 'document', signal } = {}) {
    requirePageLayout();
    if (!['document', 'selection'].includes(scope)) throw new Error('Unknown page layout scope.');
    let result;
    await Word.run(async (context) => { result = (await snapshot(context, { selection: scope === 'selection', signal })).allowed.map(publicRecord); });
    return result;
}

export async function preparePageLayoutProposal(deps, { instruction, scope = 'document', onToken, onReasoning, signal } = {}) {
    requirePageLayout(); check(signal);
    if (!['document', 'selection'].includes(scope)) throw new Error('Unknown page layout scope.');
    let captured;
    await Word.run(async (context) => { captured = await snapshot(context, { selection: scope === 'selection', signal }); });
    const config = getActiveBackendConfig(deps.appState);
    const ids = captured.allowed.map((r) => r.id);
    const raw = explicitLineNumberRemoval(instruction) ? JSON.stringify([{ sectionIds: ids, patch: { lineNumbering: { isActive: false } } }])
        : await deps.sendActionRequest(deps, config, buildPageLayoutPrompt(instruction, captured.allowed.map(publicRecord)), { onToken, onReasoning, signal });
    check(signal);
    const parsed = parsePageLayoutOps(raw, ids);
    const ops = parsed.filter((op) => !satisfied(captured.records.find((r) => r.id === op.sectionId), op.patch)).map((op) => {
        const before = captured.records.find((r) => r.id === op.sectionId);
        if (before.protectedForForms) throw new Error(`Section ${op.sectionId} is protected for forms. No layout changes were prepared.`);
        requireReadable(before, op.patch);
        validateGeometry(before, op.patch);
        return { ...op, label: describeLayoutChange(op.sectionId, before, op.patch) };
    });
    const anchor = {};
    proposals.set(anchor, { baseline: captured.records, unreadable: captured.unreadable, count: captured.records.length,
        signatures: new Set(ops.map(signature)), attempted: false, discarded: false });
    deps.log(`Prepared page layout for ${ops.length} section(s); ${scope === 'selection' ? 'whole sections containing the selection' : 'document scope'}.`, 'info');
    return { instruction, scope, ops, anchor, model: config.model, alreadySatisfied: parsed.length > 0 && ops.length === 0,
        notice: 'These changes affect entire sections and may not be recorded as tracked changes. Body text is preserved.' };
}

export async function discardPageLayoutProposal(_deps, proposal) {
    const owned = proposals.get(proposal?.anchor);
    if (owned) owned.discarded = true;
}

export async function applyPageLayoutProposal(deps, proposal, { signal } = {}) {
    requirePageLayout(); check(signal);
    const owned = proposals.get(proposal?.anchor);
    if (!owned || owned.attempted || owned.discarded || !Array.isArray(proposal.ops) || !proposal.ops.length) throw new Error('Page layout proposal is unavailable or already attempted. Draft a new proposal.');
    const ops = proposal.ops.map((op) => ({ sectionId: op.sectionId, patch: validatePageLayoutPatch(op.patch) }));
    if (new Set(ops.map((o) => o.sectionId)).size !== ops.length || ops.some((op) => !owned.signatures.has(signature(op)))) throw new Error('Page layout proposal changed after review.');
    let mutationAttempted = false;
    try {
        await Word.run(async (context) => {
            const unreadable = new Set(owned.unreadable);
            const current = await snapshot(context, { signal, unreadable });
            if (snapshotKey(current.records, unreadable) !== snapshotKey(owned.baseline, unreadable)) throw new Error('Section content or layout changed after preparation. Draft a new proposal.');
            for (const op of ops) validateGeometry(current.records.find((r) => r.id === op.sectionId), op.patch);
            check(signal);
            owned.attempted = true;
            for (const op of ops) {
                const setup = current.handles[Number(op.sectionId.slice(1)) - 1].pageSetup;
                mutationAttempted = true;
                if (op.patch.pageSetup) setup.set(op.patch.pageSetup);
                if (op.patch.lineNumbering) setup.lineNumbering.set(op.patch.lineNumbering);
                const columns = op.patch.columns;
                if (columns) {
                    if (columns.count !== undefined) setup.textColumns.setCount(columns.count);
                    if (columns.evenlySpaced !== undefined) setup.textColumns.setIsEvenlySpaced(columns.evenlySpaced);
                    if (columns.lineBetween !== undefined) setup.textColumns.setHasLineBetween(columns.lineBetween);
                    if (columns.flowDirection !== undefined) setup.textColumns.setFlowDirection(columns.flowDirection);
                }
            }
            await context.sync(); check(signal);
            for (const op of ops.filter((o) => o.patch.columns?.items)) {
                const columns = current.handles[Number(op.sectionId.slice(1)) - 1].pageSetup.textColumns;
                columns.load('items');
                await context.sync(); check(signal);
                if (columns.items.length !== op.patch.columns.items.length) throw new Error('Word did not create the requested column count.');
                op.patch.columns.items.forEach((item, i) => Object.assign(columns.items[i], item));
            }
            if (ops.some((op) => op.patch.columns?.items)) { await context.sync(); check(signal); }
            const after = await snapshot(context, { signal, unreadable });
            if (after.records.length !== owned.count || after.records.some((r, i) => r.bodyFingerprint !== current.records[i].bodyFingerprint)) throw new Error('Section body structure changed during page layout.');
            const targets = new Set(ops.map((op) => op.sectionId));
            if (after.records.some((r, i) => !targets.has(r.id) && snapshotKey([r], unreadable) !== snapshotKey([current.records[i]], unreadable))) {
                throw new Error('Word also changed layout outside the reviewed sections.');
            }
            for (const op of ops) if (!satisfied(after.records.find((r) => r.id === op.sectionId), op.patch)) throw new Error(`Word read-back did not confirm page layout for ${op.sectionId}.`);
        });
    } catch (error) {
        if (mutationAttempted) {
            const uncertain = new Error(`Page layout may be partially applied. Inspect Word before retrying. ${error.message}`);
            uncertain.mutationAttempted = true;
            throw uncertain;
        }
        throw error;
    }
    deps.log(`Page layout read-back verified ${ops.length} section(s).`, 'success');
    return { applied: true, verified: true, appliedRanges: ops.length, verifiedSections: ops.length };
}
