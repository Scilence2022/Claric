/** Addressable, Word-free document draft. Original block identities never move. */
export const DOCUMENT_EDIT_LIMITS = Object.freeze({
    blocks: 20000, paragraphChars: 24000, changes: 24, patchChars: 48000, readChars: 32000,
});

function fail(message) { throw new Error(message); }
/** @param {any} value @param {string} label @param {number} [max] */
function text(value, label, max = DOCUMENT_EDIT_LIMITS.paragraphChars) {
    if (typeof value !== 'string' || !value.trim() || value.length > max) fail(`Invalid ${label}: expected non-empty text up to ${max} characters.`);
    return value;
}
function integer(value, fallback, max) {
    if (value === undefined) return fallback;
    if (!Number.isInteger(value) || value < 0 || value > max) fail('Invalid paging range.');
    return value;
}
function insertedFormat(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)
        || !Object.keys(value).length || Object.keys(value).some((key) => !['bold', 'italic'].includes(key))
        || Object.values(value).some((item) => typeof item !== 'boolean')) fail('New paragraph format supports only explicit bold and italic booleans.');
    return { ...value };
}

/**
 * @param {{id: string, blocks: Array<{id: string, text: string, [key: string]: any}>}} snapshot
 */
export function createDocumentModel(snapshot) {
    if (!snapshot?.id || !Array.isArray(snapshot.blocks) || !snapshot.blocks.length
        || snapshot.blocks.length > DOCUMENT_EDIT_LIMITS.blocks) fail('Document snapshot is empty or too large.');
    /** @type {Array<{id: string, text: string, original: boolean, [key: string]: any}>} */
    const original = snapshot.blocks.map((b) => ({ ...b, readOnly: !!b.readOnly || !!b.headingLevel, original: true }));
    const originals = new Map(original.map((b) => [b.id, b]));
    if (originals.size !== original.length || original.some((b) => !b.id || typeof b.text !== 'string')) fail('Invalid document block identities.');
    let draft = original.map((b) => ({ ...b }));
    let revision = 0;
    let sequence = 0;
    /** @type {{goal: string, requirements: string[], targetIds: string[]}|null} */
    let contract = null;
    const readCoverage = new Map();

    function block(id, list = draft) {
        const found = list.find((b) => b.id === id);
        if (!found) fail(`Unknown block ${id}. Read the outline or search first.`);
        return found;
    }
    function readRequired(id) {
        if (id && originals.has(id) && !readCoverage.get(id)?.complete) fail(`Read the complete block ${id} before editing or choosing this gap.`);
    }
    function permitted(id) {
        if (contract?.targetIds.length && id && !contract.targetIds.includes(id)) fail(`Block ${id} is outside the declared write scope.`);
    }
    function describe(b) {
        return { id: b.id, headingLevel: b.headingLevel || 0, section: b.section || '',
            readOnly: !!b.readOnly, inTable: !!b.inTable, chars: b.text.length, preview: b.text.slice(0, 180) };
    }
    function outline(args = {}) {
        const offset = integer(args.offset, 0, draft.length);
        const limit = integer(args.limit, 40, 100);
        return { snapshotId: snapshot.id, revision, total: draft.length,
            blocks: draft.slice(offset, offset + limit).map(describe), nextOffset: offset + limit < draft.length ? offset + limit : null };
    }
    function search(args) {
        const query = text(args.query, 'search query', 200).toLocaleLowerCase();
        const matches = draft.filter((b) => b.text.toLocaleLowerCase().includes(query));
        const offset = integer(args.offset, 0, matches.length);
        return { total: matches.length, blocks: matches.slice(offset, offset + 20).map((b) => {
            const start = Math.max(0, b.text.toLocaleLowerCase().indexOf(query) - 80);
            return { ...describe(b), preview: b.text.slice(start, start + 300) };
        }), nextOffset: offset + 20 < matches.length ? offset + 20 : null };
    }
    function read(args) {
        if (!Array.isArray(args.ids) || !args.ids.length || args.ids.length > 12) fail('Read 1–12 block IDs at a time.');
        const offset = integer(args.offset, 0, Number.MAX_SAFE_INTEGER);
        const limit = integer(args.limit, 8000, DOCUMENT_EDIT_LIMITS.paragraphChars);
        let remaining = DOCUMENT_EDIT_LIMITS.readChars;
        return { revision, blocks: args.ids.map((id) => {
            const b = block(id);
            const count = Math.min(limit, remaining);
            const excerpt = b.text.slice(offset, offset + count);
            remaining -= excerpt.length;
            if (b.original && b.text === originals.get(id).text) {
                const coverage = readCoverage.get(id) || { ranges: [], complete: false };
                coverage.ranges.push([Math.min(offset, b.text.length), Math.min(offset + excerpt.length, b.text.length)]);
                coverage.ranges.sort((a, b) => a[0] - b[0]);
                let end = 0;
                for (const [a, z] of coverage.ranges) { if (a > end) break; end = Math.max(end, z); }
                coverage.complete = end === b.text.length;
                readCoverage.set(id, coverage);
            }
            const index = draft.indexOf(b);
            return { ...describe(b), text: excerpt, offset,
                nextOffset: offset + excerpt.length < b.text.length ? offset + excerpt.length : null,
                previousId: draft[index - 1]?.id || null, nextId: draft[index + 1]?.id || null };
        }) };
    }
    function setContract(args) {
        if (revision) fail('The edit contract cannot change after drafting. Discard the draft before changing the task.');
        const goal = text(args.goal, 'goal', 2000);
        if (!Array.isArray(args.requirements) || !args.requirements.length || args.requirements.length > 10) fail('List 1–10 requirements, including placement, scope and preservation constraints.');
        const requirements = args.requirements.map((r) => text(r, 'requirement', 1000));
        const targetIds = args.targetIds === undefined ? [] : args.targetIds;
        if (!Array.isArray(targetIds) || targetIds.length > 100) fail('Invalid targetIds.');
        targetIds.forEach((id) => { if (!originals.has(id)) fail(`Unknown original target ${id}.`); });
        contract = { goal, requirements, targetIds: [...new Set(targetIds)] };
        return { contract, revision };
    }
    function compile(list = draft) {
        const changes = [];
        for (let i = 0; i < list.length; i++) {
            const b = list[i];
            if (b.original) {
                if (b.text !== originals.get(b.id).text) changes.push({ id: `replace-${b.id}`, kind: 'replace',
                    blockId: b.id, before: originals.get(b.id).text, after: b.text, reason: b.reason || '' });
                continue;
            }
            const afterId = list[i - 1]?.id || null;
            const inserted = [b];
            while (list[i + 1] && !list[i + 1].original) inserted.push(list[++i]);
            const beforeId = list[i + 1]?.id || null;
            const formats = inserted.map((p) => p.format || null);
            changes.push({ id: `insert-${inserted[0].id}`, kind: 'insert', afterId, beforeId,
                paragraphs: inserted.map((p) => p.text), ...(formats.some(Boolean) ? { paragraphFormats: formats } : {}),
                reason: inserted.map((p) => p.reason || '').filter(Boolean).join('; ') });
        }
        if (changes.length > DOCUMENT_EDIT_LIMITS.changes || JSON.stringify(changes).length > DOCUMENT_EDIT_LIMITS.patchChars) fail('Patch is too large. Reduce the draft to focused changes.');
        return { snapshotId: snapshot.id, revision, changes };
    }
    function stage(args) {
        if (!contract) fail('Call set_edit_contract before editing.');
        if (!Array.isArray(args.operations) || !args.operations.length || args.operations.length > 12) fail('Stage 1–12 operations at a time.');
        const next = draft.map((b) => ({ ...b }));
        let nextSequence = sequence;
        for (const op of args.operations) {
            if (!op || typeof op !== 'object') fail('Invalid operation.');
            if (op.kind === 'insert') {
                if ((op.afterId == null) === (op.beforeId == null)) fail('Specify exactly one of afterId or beforeId. Use a real block ID, never a position default.');
                const anchor = block(op.afterId ?? op.beforeId, next);
                let index = next.indexOf(anchor) + (op.afterId != null ? 1 : 0);
                // Original neighbors define the address even when editing an earlier insertion.
                const left = next.slice(0, index).reverse().find((b) => b.original);
                const right = next.slice(index).find((b) => b.original);
                readRequired(left?.id); readRequired(right?.id);
                const owner = op.afterId != null ? (left || right) : (right || left);
                permitted(owner?.id);
                if ((!left || left.inTable) && (!right || right.inTable)) fail('Cannot insert prose inside a table. Choose a body paragraph.');
                if (!Array.isArray(op.paragraphs) || !op.paragraphs.length || op.paragraphs.length > 12) fail('Insert 1–12 paragraphs.');
                const reason = text(op.reason, 'placement reason', 1000);
                for (const value of op.paragraphs) {
                    const content = text(value, 'paragraph');
                    if (/[\r\n]/.test(content)) fail('Each paragraph must be a single block; use multiple array entries.');
                    do { nextSequence++; } while (originals.has(`draft-${nextSequence}`));
                    next.splice(index++, 0, { id: `draft-${nextSequence}`, text: content, original: false,
                        reason, headingLevel: 0, section: anchor.section || '', readOnly: false, inTable: false });
                }
            } else if (op.kind === 'replace') {
                const target = block(op.blockId, next);
                readRequired(target.id);
                if (target.original) permitted(target.id);
                if (target.readOnly) fail(`Block ${target.id} contains protected structure. Choose a plain body paragraph.`);
                const content = text(op.text, 'replacement');
                if (/[\r\n]/.test(content)) fail('A replacement must remain one paragraph. Insert additional paragraphs separately.');
                target.text = content;
                target.reason = text(op.reason, 'edit reason', 1000);
            } else if (op.kind === 'format_new') {
                const target = block(op.blockId, next);
                if (target.original) fail('format_new can only target a newly inserted draft paragraph.');
                target.format = insertedFormat(op.format);
                target.reason = [target.reason, text(op.reason, 'format reason', 1000)].filter(Boolean).join('; ');
            } else if (op.kind === 'discard') {
                const target = block(op.blockId, next);
                if (target.original) Object.assign(target, originals.get(target.id), { reason: '' });
                else next.splice(next.indexOf(target), 1);
            } else fail(`Unsupported operation ${op.kind}. Use insert, replace, format_new or discard.`);
        }
        compile(next); // Commit only after the entire batch validates.
        draft = next;
        sequence = nextSequence;
        revision++;
        return { revision, changes: compile().changes, message: 'Draft updated; Word has not been changed. Read the draft and validate it before finishing.' };
    }
    function preview() {
        const patch = compile();
        const ids = new Set();
        for (const change of patch.changes) {
            for (const id of [change.blockId, 'afterId' in change ? change.afterId : null,
                'beforeId' in change ? change.beforeId : null].filter(Boolean)) {
                const i = original.findIndex((b) => b.id === id);
                for (const neighbor of original.slice(Math.max(0, i - 1), i + 2)) ids.add(neighbor.id);
            }
        }
        const before = original.filter((b) => ids.has(b.id)).map((b) => ({ id: b.id, text: b.text, section: b.section || '' }));
        const after = draft.filter((b) => !b.original || ids.has(b.id)).map((b) => ({ id: b.id, text: b.text,
            section: b.section || '', ...(b.format ? { format: b.format } : {}) }));
        const result = { contract, ...patch, before, after };
        if (JSON.stringify(result).length > 60000) fail('Review context is too large. Reduce the draft to fewer paragraphs.');
        return result;
    }
    return { outline, search, read, setContract, stage, compile, preview,
        get revision() { return revision; }, get contract() { return contract; } };
}
