import { rangeStructureFingerprint } from '../lib/ooxml-fingerprint.js';
import { checkOperationSignal, discardFormatProposal } from './word-actions.js';
export { inferTableCreationSpec, buildTableCreationPrompt, parseTableCreationResponse, validateTableCreationSpec } from '../lib/table-ops.js';

let tableAnchorSequence = 0;
export async function captureTableAnchor(deps, signal) {
    const anchor = {};
    try {
        await Word.run(async (context) => {
            checkOperationSignal(signal);
            const selection = context.document.getSelection();
            if (typeof selection.insertBookmark !== 'function'
                || typeof context.document.getBookmarkRangeOrNullObject !== 'function') return;
            anchor.bookmark = `_claric_tbl_${Date.now().toString(36)}_${++tableAnchorSequence}`;
            selection.insertBookmark(anchor.bookmark);
            await context.sync();
            checkOperationSignal(signal);
            const range = context.document.getBookmarkRangeOrNullObject(anchor.bookmark);
            range.load('isNullObject,text');
            await context.sync();
            checkOperationSignal(signal);
            if (range.isNullObject || typeof range.getOoxml !== 'function') return;
            const xml = range.getOoxml();
            await context.sync();
            checkOperationSignal(signal);
            anchor.fingerprint = rangeStructureFingerprint(xml.value);
            anchor.text = range.text;
        });
        if (anchor.fingerprint && typeof anchor.text === 'string') return anchor;
        await discardTableProposal(deps, { anchor });
        return null;
    } catch (error) {
        await discardTableProposal(deps, { anchor });
        if (signal?.aborted || error.name === 'AbortError') throw error;
        deps.log(`Could not capture table placement: ${error.message}`, 'warning');
        return null;
    }
}

export async function discardTableProposal(deps, proposal) {
    try { await discardFormatProposal(deps, proposal); }
    catch (error) { deps.log(`Table placement cleanup failed: ${error.message}`, 'warning'); }
}

export async function resolveTableAnchor(context, anchor, signal) {
    if (!anchor?.bookmark || !anchor.fingerprint || anchor.cleaned) throw new Error('The table placement is not anchored. Draft a new proposal.');
    const range = context.document.getBookmarkRangeOrNullObject(anchor.bookmark);
    range.load('isNullObject,text');
    await context.sync();
    checkOperationSignal(signal);
    if (range.isNullObject || range.text !== anchor.text) throw new Error('The reviewed table placement changed. Draft a new proposal.');
    const xml = range.getOoxml();
    await context.sync();
    checkOperationSignal(signal);
    if (rangeStructureFingerprint(xml.value) !== anchor.fingerprint) throw new Error('The reviewed table placement changed. Draft a new proposal.');
    return range;
}
