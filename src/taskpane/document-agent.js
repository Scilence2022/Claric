import {
    readCursorContext, prepareSelectionAmendment, applySelectionAmendment,
    prepareDocumentAppend, applyDocumentAppend,
    prepareFormatProposal, applyFormatProposal, discardFormatProposal,
    prepareTableProposal, applyTableProposal,
} from './word-actions.js';
import { extractDocumentStructured, estimateTokenCount } from '../lib/comment-extractor.js';
import { describeFormatOp } from '../lib/format-ops.js';

const DEFAULT_MAX_CHARS = 12000;
const DEFAULT_MAX_TOKENS = 3000;
const MAX_PROPOSAL_TEXT = 4000;
const MAX_PREPARED_TASKS = 50;

function fail(code, message) { return Object.assign(new Error(message), { code }); }
function aborted(signal) { if (signal?.aborted) throw new DOMException('Operation cancelled', 'AbortError'); }
function prefix(text, length) {
    let end = Math.min(length, text.length);
    if (end > 0 && /[\uD800-\uDBFF]/.test(text[end - 1])) end -= 1;
    return text.slice(0, end);
}
function limit(value, fallback) {
    return Number.isSafeInteger(value) && value > 0 ? Math.min(value, fallback) : fallback;
}
function tablePreview(rows, maxChars) {
    const lines = rows.map((row) => `| ${row.map((cell) => String(cell ?? '').replace(/\|/g, '\\|').replace(/\s+/g, ' ').trim()).join(' | ')} |`);
    return prefix(lines.join('\n'), maxChars);
}
function boundedText(text, maxChars, maxTokens) {
    let result = prefix(text, maxChars);
    if (estimateTokenCount(result) > maxTokens) {
        let low = 0;
        let high = result.length;
        while (low < high) {
            const mid = Math.ceil((low + high) / 2);
            if (estimateTokenCount(prefix(result, mid)) <= maxTokens) low = mid;
            else high = mid - 1;
        }
        result = prefix(result, low);
    }
    return result;
}
export async function revisionFor(value, cryptoImpl = globalThis.crypto) {
    if (!cryptoImpl?.subtle) throw fail('REVISION_UNAVAILABLE', 'Secure document fingerprinting is unavailable');
    const bytes = new TextEncoder().encode(typeof value === 'string' ? value : JSON.stringify(value));
    const digest = await cryptoImpl.subtle.digest('SHA-256', bytes);
    return `revision-${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`;
}
async function localSelection() {
    if (typeof Word === 'undefined') throw fail('WORD_UNAVAILABLE', 'Word is unavailable');
    return Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load('text');
        await context.sync();
        return { text: selection.text || '' };
    });
}

export function createDocumentAgent({ identity, actions = {}, appState = {}, log = () => {}, clock = Date.now, cryptoImpl = globalThis.crypto } = {}) {
    if (!identity || ['workspaceId', 'documentId', 'instanceId'].some((key) => typeof identity[key] !== 'string' || !identity[key])) {
        throw fail('INVALID_IDENTITY', 'Document agent requires a complete local identity');
    }
    const owner = { ...identity };
    const readSelection = actions.readSelectionContent || localSelection;
    const readCursor = actions.readCursorContext || readCursorContext;
    const readDocument = actions.extractDocumentStructured || extractDocumentStructured;
    const prepareSelection = actions.prepareSelectionAmendment || prepareSelectionAmendment;
    const applySelection = actions.applySelectionAmendment || applySelectionAmendment;
    const prepareAppend = actions.prepareDocumentAppend || prepareDocumentAppend;
    const applyAppend = actions.applyDocumentAppend || applyDocumentAppend;
    const prepareFormat = actions.prepareFormatProposal || prepareFormatProposal;
    const applyFormat = actions.applyFormatProposal || applyFormatProposal;
    const discardFormat = actions.discardFormatProposal || discardFormatProposal;
    const prepareTable = actions.prepareTableProposal || prepareTableProposal;
    const applyTable = actions.applyTableProposal || applyTableProposal;
    const preparedTasks = new Map();
    const deps = { appState, log };
    let disposed = false;
    let busy = false;

    const KINDS = {
        edit: { scope: 'selection', itemId: 'text-1', label: 'Selected passage', title: 'Proposed passage edit' },
        append: { scope: 'document', itemId: 'text-1', label: 'New content', title: 'Proposed appended content' },
        format: { scope: 'selection', itemId: 'format-1', label: 'Formatting changes', title: 'Proposed formatting' },
        table: { scope: 'selection', itemId: 'table-1', label: 'New table', title: 'Proposed table insertion' },
    };

    function check(signal) {
        aborted(signal);
        if (disposed) throw fail('AGENT_STOPPED', 'Document agent has stopped');
    }
    async function read(scope, signal) {
        check(signal);
        let raw;
        if (scope === 'selection') raw = await readSelection(deps);
        else if (scope === 'cursor') raw = await readCursor(deps);
        else if (scope === 'document' || scope === 'outline') raw = await readDocument({ richness: 'headings' });
        else throw fail('UNSUPPORTED_SCOPE', `Unsupported context scope: ${scope}`);
        check(signal);
        if (raw === null || raw === undefined) throw fail('CONTEXT_UNAVAILABLE', 'Requested context is unavailable');
        const text = typeof raw === 'string' ? raw : raw.text ?? raw.contextText;
        if (typeof text !== 'string') throw fail('INVALID_CONTEXT', 'Document reader returned no text');
        return text;
    }
    async function getDocumentRevision(scope = 'selection', { signal } = {}) {
        return revisionFor({ documentId: owner.documentId, scope, text: await read(scope, signal) }, cryptoImpl);
    }
    function getCapabilities() {
        return { readContext: true, readSelection: true, readCursor: true, readDocument: true, readOutline: true,
            prepareProposal: true, applyProposal: true, taskTypes: Object.keys(KINDS) };
    }
    async function readContext(request = {}, { signal } = {}) {
        const scope = request.scope || 'selection';
        const text = await read(scope, signal);
        const selected = scope === 'outline' ? text.split('\n').filter((line) => /^#{1,9}\s/.test(line)).join('\n') : text;
        const maxChars = limit(request.maxChars, DEFAULT_MAX_CHARS);
        const maxTokens = limit(request.maxTokens, DEFAULT_MAX_TOKENS);
        const excerpt = boundedText(selected, maxChars, maxTokens);
        const documentRevision = await revisionFor({ documentId: owner.documentId, scope, text }, cryptoImpl);
        check(signal);
        return {
            snapshotId: `snapshot-${cryptoImpl.randomUUID()}`, sourceDocumentId: owner.documentId,
            sourceInstanceId: owner.instanceId, scope, generatedAt: new Date(clock()).toISOString(),
            documentRevision, text: excerpt, outline: [], tables: [], citations: [],
            truncation: { text: excerpt.length < selected.length, maxChars, maxTokens },
        };
    }
    async function exclusive(operation, signal) {
        check(signal);
        if (busy || appState.isProcessing || appState.isProcessingDoc || appState.isProcessingSummary || appState.chatController || appState.processDocController) {
            throw fail('DOCUMENT_BUSY', 'Another operation is using this document');
        }
        busy = true;
        appState.isProcessingDoc = true;
        try { return await operation(); }
        finally { busy = false; appState.isProcessingDoc = false; }
    }
    async function prepareTask(task = {}, { signal = task.signal } = {}) {
        return exclusive(async () => {
            const instruction = typeof task.instruction === 'string' ? task.instruction.trim() : '';
            if (!instruction || typeof task.taskId !== 'string' || !task.taskId) throw fail('INVALID_TASK', 'Task identity and instruction are required');
            if (preparedTasks.has(task.taskId)) throw fail('DUPLICATE_TASK', 'Task was already prepared');
            if (preparedTasks.size >= MAX_PREPARED_TASKS) throw fail('CAPACITY', 'Review or discard pending proposals first');
            const kind = task.taskType || task.type;
            const config = KINDS[kind];
            if (!config) throw fail('UNSUPPORTED_TASK', `Unsupported remote task type: ${kind}`);
            const scope = config.scope;
            if (task.scope && task.scope !== scope) throw fail('UNSUPPORTED_SCOPE', `Task requires ${scope} scope`);
            const original = await read(scope, signal);
            if (['edit', 'format'].includes(kind) && !original.trim()) throw fail('EMPTY_SELECTION', 'Select the target passage before requesting an edit');
            const baseRevision = await revisionFor({ documentId: owner.documentId, scope, text: original }, cryptoImpl);
            let proposal;
            let before = '';
            let after = '';
            if (kind === 'edit') {
                proposal = await prepareSelection(deps, { promptTemplate: instruction, signal });
                check(signal);
                if (!proposal || proposal.tablePatch || proposal.mixedTable) throw fail('UNSUPPORTED_PROPOSAL', 'This remote task requires the local structured-edit workflow');
                before = proposal.selectionText;
                after = proposal.amendedText;
            } else if (kind === 'append') {
                proposal = await prepareAppend(deps, { instruction, signal });
                check(signal);
                after = proposal?.generatedText;
            } else if (kind === 'format') {
                proposal = await prepareFormat(deps, { instruction, scope: 'selection', selectionText: original, signal });
                check(signal);
                if (!proposal || !Array.isArray(proposal.ops) || !proposal.ops.length) throw fail('NO_CHANGES', 'The model proposed no formatting changes');
                before = original;
                after = proposal.ops.map((op) => describeFormatOp(op)).join('\n');
            } else {
                proposal = await prepareTable(deps, { instruction, signal });
                check(signal);
                if (!proposal || !Array.isArray(proposal.spec?.rows) || !proposal.spec.rows.length || !Array.isArray(proposal.spec.rows[0])) throw fail('NO_CHANGES', 'The model proposed no table content');
                after = tablePreview(proposal.spec.rows, MAX_PROPOSAL_TEXT);
            }
            if (typeof before !== 'string' || typeof after !== 'string' || !after.trim() || before === after) throw fail('NO_CHANGES', 'The model proposed no changes');
            if (before.length > MAX_PROPOSAL_TEXT || after.length > MAX_PROPOSAL_TEXT) throw fail('PROPOSAL_TOO_LARGE', 'This proposal exceeds the remote review limit; use a smaller passage');
            if (await getDocumentRevision(scope, { signal }) !== baseRevision) throw fail('STALE', 'Document changed while preparing the proposal');
            const expiresAt = clock() + 15 * 60 * 1000;
            preparedTasks.set(task.taskId, { proposal, baseRevision, kind, scope, itemId: config.itemId, expiresAt, attempted: false });
            return { taskId: task.taskId, kind, scope, title: config.title,
                summary: instruction.slice(0, 240), baseRevision, expiresAt,
                items: [{ id: config.itemId, label: config.label, before, after }] };
        }, signal);
    }
    async function applyProposal(selectedItemIds, record = {}, { signal } = {}) {
        return exclusive(async () => {
            if (['workspaceId', 'documentId', 'instanceId'].some((key) => record.target?.[key] !== owner[key])) throw fail('TARGET_MISMATCH', 'Proposal belongs to another document instance');
            const prepared = preparedTasks.get(record.taskId);
            if (!prepared) throw fail('PROPOSAL_RUNTIME_MISSING', 'Target-local prepared proposal is unavailable');
            if (prepared.attempted) throw fail('ALREADY_ATTEMPTED', 'A write was already attempted; inspect the document before generating a new proposal');
            if (prepared.expiresAt <= clock()) throw fail('EXPIRED', 'Proposal has expired');
            if (!Array.isArray(selectedItemIds) || selectedItemIds.length !== 1 || String(selectedItemIds[0]) !== prepared.itemId) throw fail('INVALID_ITEMS', 'Select the prepared change');
            if (record.baseRevision !== prepared.baseRevision || await getDocumentRevision(prepared.scope, { signal }) !== prepared.baseRevision) throw fail('STALE', 'Document changed since preparation');
            check(signal);
            prepared.attempted = true;
            const apply = { edit: applySelection, append: applyAppend, format: applyFormat, table: applyTable }[prepared.kind];
            const result = await apply(deps, prepared.proposal);
            if (result?.skipped || result?.interrupted || result?.errors?.length) throw fail('WRITE_NOT_COMPLETED', result.reason || 'Word did not finish applying the proposal');
            return { appliedItemIds: [prepared.itemId], documentRevision: await getDocumentRevision(prepared.scope), result };
        }, signal);
    }
    async function discardPrepared(taskId) {
        const prepared = preparedTasks.get(taskId);
        if (!prepared) return;
        preparedTasks.delete(taskId);
        // Format preparation anchors its scope with a bookmark; remove it so a
        // rejected or expired proposal leaves no markup behind.
        if (prepared.kind === 'format' && !prepared.attempted) {
            try { await discardFormat(deps, prepared.proposal); }
            catch (error) { log(`Remote format cleanup failed: ${error.message}`, 'warning'); }
        }
    }
    function discardProposal(record) { void discardPrepared(typeof record === 'string' ? record : record?.taskId); }
    function dispose() {
        disposed = true;
        for (const taskId of [...preparedTasks.keys()]) void discardPrepared(taskId);
    }
    return Object.freeze({ getCapabilities, getDocumentRevision, readContext, prepareTask, applyProposal, discardProposal, dispose });
}
export { DEFAULT_MAX_CHARS, DEFAULT_MAX_TOKENS };
