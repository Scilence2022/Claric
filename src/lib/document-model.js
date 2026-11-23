/** Addressable, Word-free document draft with stable original block identities. */
import { resolveDocumentMove } from './document-move.js';

export const DOCUMENT_EDIT_LIMITS = Object.freeze({
    blocks: 20000, paragraphChars: 24000, changes: 24, patchChars: 48000, readChars: 32000,
    outlineBlocks: 100, readBlocks: 12,
});

/** Hard cap on the serialized review preview, and the default unchanged-context excerpt per side. */
const REVIEW_CONTEXT_CHARS = 60000;
const REVIEW_EXCERPT_CHARS = 1500;

function fail(message) { throw new Error(message); }
/** @param {any} value @param {string} label @param {number} [max] */
function text(value, label, max = DOCUMENT_EDIT_LIMITS.paragraphChars) {
    if (typeof value !== 'string' || !value.trim() || value.length > max) fail(`Invalid ${label}: expected non-empty text up to ${max} characters.`);
    return value;
}
function integer(value, fallback, max, label = 'paging value', min = 0) {
    if (value === undefined) return fallback;
    if (!Number.isInteger(value) || value < min || value > max) fail(`Invalid ${label}: use an integer from ${min} to ${max}.`);
    return value;
}
function insertedFormat(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value)
        || !Object.keys(value).length || Object.keys(value).some((key) => !['bold', 'italic'].includes(key))
        || Object.values(value).some((item) => typeof item !== 'boolean')) fail('New paragraph format supports only explicit bold and italic booleans.');
    return { ...value };
}

/**
 * Bound an unchanged context block to the excerpt(s) facing the edit: its end
 * when a change follows, its start when a change precedes. The omitted span is
 * named inline and in `elided` so it cannot be mistaken for a deletion.
 * @param {string} value @param {number} limit @param {{head: boolean, tail: boolean}} sides
 */
function elide(value, limit, { head, tail }) {
    const keepHead = head || !tail ? limit : 0;
    const keepTail = tail ? limit : 0;
    if (value.length <= keepHead + keepTail + 200) return {};
    const start = keepHead;
    const end = value.length - keepTail;
    return { text: `${value.slice(0, start)}[… unchanged text elided: characters ${start}–${end} of ${value.length} …]${value.slice(end)}`,
        elided: { start, end, chars: value.length } };
}

/**
 * @param {{id: string, authorizedIds?: string[], blocks: Array<{id: string, text: string, [key: string]: any}>}} snapshot
 */
export function createDocumentModel(snapshot) {
    if (!snapshot?.id || !Array.isArray(snapshot.blocks) || !snapshot.blocks.length
        || snapshot.blocks.length > DOCUMENT_EDIT_LIMITS.blocks) fail('Document snapshot is empty or too large.');
    /** @type {Array<{id: string, text: string, original: boolean, [key: string]: any}>} */
    const original = snapshot.blocks.map((b) => ({ ...b,
        readOnly: !!b.readOnly || !!b.headingLevel || !!b.structureUnavailable || !!b.inTable || !!b.isListItem, original: true }));
    const originals = new Map(original.map((b) => [b.id, b]));
    const originalPositions = new Map(original.map((b, index) => [b.id, index]));
    const authorized = snapshot.authorizedIds && new Set(snapshot.authorizedIds);
    if (originals.size !== original.length || original.some((b) => !b.id || typeof b.text !== 'string')) fail('Invalid document block identities.');
    let draft = original.map((b) => ({ ...b }));
    let moves = [];
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
        if (id && !inScope(id)) fail(`Block ${id} is outside the authorized write scope.`);
    }
    function inScope(id) { return (!authorized || authorized.has(id)) && (!contract?.targetIds.length || contract.targetIds.includes(id)); }
    function describe(b) {
        const index = draft.indexOf(b);
        function canInsert(at, after) {
            const left = draft.slice(0, at).reverse().find((item) => item.original);
            const right = draft.slice(at).find((item) => item.original);
            const owner = after ? (left || right) : (right || left);
            return !left?.structureUnavailable && !right?.structureUnavailable
                && !((!left || left.inTable) && (!right || right.inTable))
                && inScope(owner?.id) && (!authorized || [left, right].every((b) => !b || authorized.has(b.id)));
        }
        const reasons = b.readOnlyReasons || [b.structureUnavailable && 'structure_unavailable',
            b.headingLevel && 'heading', b.inTable && 'table', b.isListItem && 'list',
            b.readOnly && !b.structureUnavailable && !b.headingLevel && !b.inTable && !b.isListItem && 'protected_structure'].filter(Boolean);
        return { id: b.id, headingLevel: b.headingLevel || 0, section: b.section || '',
            hasRevisions: !!b.hasRevisions, textVersion: 'current',
            readOnly: !!b.readOnly, readOnlyReasons: reasons, inTable: !!b.inTable,
            canReplace: !b.readOnly && (!b.original || inScope(b.id)),
            canMove: b.canMove === true && inScope(b.id) && !moves.some((move) => move.blockIds.includes(b.id)),
            canAnchorMove: b.canAnchorMove === true && inScope(b.id),
            canInsertBefore: canInsert(index, false), canInsertAfter: canInsert(index + 1, true),
            chars: b.text.length, preview: b.text.slice(0, 180) };
    }
    function outline(args = {}) {
        if (args.headingsOnly !== undefined && typeof args.headingsOnly !== 'boolean') fail('headingsOnly must be a boolean.');
        const list = args.headingsOnly ? draft.filter((b) => b.headingLevel) : draft;
        const offset = integer(args.offset, 0, list.length, 'outline offset');
        const limit = integer(args.limit, 40, DOCUMENT_EDIT_LIMITS.outlineBlocks, 'outline limit', 1);
        return { snapshotId: snapshot.id, revision, total: list.length,
            blocks: list.slice(offset, offset + limit).map(describe), nextOffset: offset + limit < list.length ? offset + limit : null };
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
        if (!Array.isArray(args.ids) || !args.ids.length || args.ids.length > DOCUMENT_EDIT_LIMITS.readBlocks) fail('Read 1–12 block IDs at a time. Split larger requests into pages of at most 12 IDs.');
        const offset = integer(args.offset, 0, Number.MAX_SAFE_INTEGER, 'read offset');
        const limit = integer(args.limit, 8000, DOCUMENT_EDIT_LIMITS.paragraphChars, 'read limit', 1);
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
            return { ...describe(b), originalText: b.original && b.text === originals.get(id).text, text: excerpt, offset,
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
    function compile(list = draft, structural = moves) {
        const changes = [...structural];
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
        const nextMoves = moves.slice();
        let nextSequence = sequence;
        for (const op of args.operations) {
            if (!op || typeof op !== 'object') fail('Invalid operation.');
            if (op.kind === 'move') {
                if ((op.afterId == null) === (op.beforeId == null)) fail('Specify one destination using beforeId or afterId.');
                const first = next.indexOf(block(op.firstId, next));
                const last = next.indexOf(block(op.lastId, next));
                const source = next.slice(first, last + 1);
                const target = block(op.afterId ?? op.beforeId, next);
                const side = op.afterId == null ? 'before' : 'after';
                const blockIds = source.map((b) => b.id);
                const move = { kind: 'move', id: `move-${op.firstId}`, blockIds, targetId: target.id, side,
                    reason: text(op.reason, 'move reason', 1000) };
                const { start, end, gap, destination } = resolveDocumentMove(next.map((b) => b.id), move);
                if (!target.canAnchorMove || source.some((b) => b.canMove !== true)) fail('Move requires verified body paragraphs without protected content or unresolved revisions, and a verified body destination.');
                if (blockIds.some((id, index) => originalPositions.get(id) !== originalPositions.get(blockIds[0]) + index)) fail('Choose a source that was also contiguous in the original Word document.');
                resolveDocumentMove(next.filter((b) => b.original).map((b) => b.id), move);
                if (nextMoves.some((item) => item.blockIds.some((id) => blockIds.includes(id)))) fail('A paragraph can be moved once per draft. Discard the latest move to revise its destination.');
                for (const id of [...blockIds, target.id]) permitted(id);
                const sourceNeighbors = [next[start - 1]?.id || null, next[end]?.id || null];
                const destinationNeighbors = [next[gap - 1]?.id || null, next[gap]?.id || null];
                for (const id of [op.firstId, op.lastId, ...sourceNeighbors, ...destinationNeighbors]) readRequired(id);
                next.splice(start, source.length);
                next.splice(destination, 0, ...source);
                nextMoves.push({ ...move, sourceNeighbors, destinationNeighbors });
            } else if (op.kind === 'discard_move') {
                const move = nextMoves.at(-1);
                if (!move || move.id !== op.moveId) fail('Discard moves starting with the latest moveId.');
                const nextId = move.sourceNeighbors[1];
                const reverse = { blockIds: move.blockIds, targetId: nextId || move.sourceNeighbors[0], side: nextId ? 'before' : 'after' };
                const { start, destination } = resolveDocumentMove(next.map((b) => b.id), reverse);
                const source = next.splice(start, move.blockIds.length);
                next.splice(destination, 0, ...source);
                nextMoves.pop();
            } else if (op.kind === 'insert') {
                if ((op.afterId == null) === (op.beforeId == null)) fail('Specify exactly one of afterId or beforeId. Use a real block ID, never a position default.');
                const anchor = block(op.afterId ?? op.beforeId, next);
                let index = next.indexOf(anchor) + (op.afterId != null ? 1 : 0);
                // Original neighbors define the address even when editing an earlier insertion.
                const left = next.slice(0, index).reverse().find((b) => b.original);
                const right = next.slice(index).find((b) => b.original);
                readRequired(left?.id); readRequired(right?.id);
                if (left?.structureUnavailable || right?.structureUnavailable) fail('Cannot insert beside a paragraph whose Word structure could not be read. Choose another location.');
                const owner = op.afterId != null ? (left || right) : (right || left);
                permitted(owner?.id);
                if (authorized && [left, right].some((b) => b && !authorized.has(b.id))) fail('Insertion gap is outside the selected paragraphs.');
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
                if (typeof op.blockId !== 'string' || !op.blockId) fail('replace requires blockId, text and reason. Use blockId for the original or draft ID, not id.');
                const target = block(op.blockId, next);
                readRequired(target.id);
                if (target.original) permitted(target.id);
                if (target.readOnly) fail(`Block ${target.id} contains protected structure. Choose a plain body paragraph.`);
                const content = text(op.text, 'replacement');
                if (/[\r\n]/.test(content)) fail('A replacement must remain one paragraph. Insert additional paragraphs separately.');
                target.text = content;
                target.reason = text(op.reason, 'edit reason', 1000);
            } else if (op.kind === 'format_new') {
                if (typeof op.blockId !== 'string' || !op.blockId) fail('format_new requires blockId, format and reason.');
                const target = block(op.blockId, next);
                if (target.original) fail('format_new can only target a newly inserted draft paragraph.');
                target.format = insertedFormat(op.format);
                target.reason = [target.reason, text(op.reason, 'format reason', 1000)].filter(Boolean).join('; ');
            } else if (op.kind === 'discard') {
                if (typeof op.blockId !== 'string' || !op.blockId) fail('discard requires blockId.');
                const target = block(op.blockId, next);
                if (target.original) Object.assign(target, originals.get(target.id), { reason: '' });
                else next.splice(next.indexOf(target), 1);
            } else fail(`Unsupported operation ${op.kind}. Use move, discard_move, insert, replace, format_new or discard.`);
        }
        if (moves.length || nextMoves.length) {
            let section = '';
            for (const b of next) {
                if (b.headingLevel) section = b.text;
                b.section = section;
            }
        }
        compile(next, nextMoves); // Commit only after the entire batch validates.
        draft = next;
        moves = nextMoves;
        sequence = nextSequence;
        revision++;
        return { revision, changes: compile().changes, message: 'Draft updated; Word has not been changed. Read the draft and validate it before finishing.' };
    }
    function preview() {
        const patch = compile();
        const ids = new Set();
        const moveEdges = new Set(moves.flatMap((move) => [move.blockIds[0], move.blockIds.at(-1),
            ...move.sourceNeighbors, ...move.destinationNeighbors].filter(Boolean)));
        for (const id of moveEdges) ids.add(id);
        for (const change of patch.changes) {
            for (const id of [change.blockId, 'afterId' in change ? change.afterId : null,
                'beforeId' in change ? change.beforeId : null].filter(Boolean)) {
                const i = original.findIndex((b) => b.id === id);
                for (const neighbor of original.slice(Math.max(0, i - 1), i + 2)) ids.add(neighbor.id);
            }
        }
        // Changed blocks are always shown in full. Unchanged neighbours are context
        // only, so a long one is cut to the end(s) facing the nearest change.
        const edited = draft.map((b) => !b.original || b.text !== originals.get(b.id).text);
        const changedIds = new Set(draft.filter((_, i) => edited[i]).map((b) => b.id));
        const firstEdited = edited.indexOf(true);
        const lastEdited = edited.lastIndexOf(true);
        const sides = new Map(draft.map((b, i) => [b.id, { head: moveEdges.has(b.id) || (firstEdited !== -1 && i >= firstEdited),
            tail: moveEdges.has(b.id) || i <= lastEdited }]));
        const view = (list, excerpt, include) => list.flatMap((b, position) => include(b) ? [{ id: b.id, text: b.text, section: b.section || '',
            ...(moves.length ? { position: position + 1 } : {}),
            ...(changedIds.has(b.id) ? {} : elide(b.text, excerpt, sides.get(b.id))), ...(b.format ? { format: b.format } : {}) }] : []);
        const full = (list) => view(list, 0, (b) => changedIds.has(b.id));
        if (JSON.stringify({ contract, ...patch, before: full(original), after: full(draft) }).length > REVIEW_CONTEXT_CHARS) {
            fail(`Review context is too large: the proposed changes alone exceed ${REVIEW_CONTEXT_CHARS} characters. Reduce the draft to fewer or shorter paragraphs.`);
        }
        // Many long neighbours can still add up; shrink their excerpts before giving up.
        for (const excerpt of [REVIEW_EXCERPT_CHARS, 600, 200, 0]) {
            const movedIds = new Set(moves.flatMap((move) => move.blockIds));
            const include = (b) => ids.has(b.id) || (excerpt === REVIEW_EXCERPT_CHARS && movedIds.has(b.id));
            const result = { contract, ...patch, before: view(original, excerpt, include),
                after: view(draft, excerpt, (b) => !b.original || include(b)) };
            if (JSON.stringify(result).length <= REVIEW_CONTEXT_CHARS) return result;
        }
        fail('Review context is too large. Reduce the draft to fewer paragraphs.');
    }
    return { outline, search, read, setContract, stage, compile, preview,
        get revision() { return revision; }, get contract() { return contract; } };
}
