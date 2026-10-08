/** Original document observations retained independently of model chat history. */
export const DOCUMENT_EVIDENCE_LIMITS = Object.freeze({ chars: 48000, pinnedBlocks: 24 });

const stopWords = new Set('a an and are as at be been by can for from has have in into is it its of on or that the their these this to was were which with'.split(' '));
function terms(value) {
    return new Set((String(value || '').toLocaleLowerCase().match(/[\p{L}\p{N}]+/gu) || [])
        .filter((word) => word.length >= 3 && !stopWords.has(word)));
}

/**
 * Keep read ranges, not repeated copies of paragraphs. Only text actually
 * observed from the original snapshot can become evidence; draft prose cannot.
 * @param {Array<{id: string, text: string}>} blocks
 */
export function createDocumentEvidenceStore(blocks) {
    const originals = new Map(blocks.map((block, index) => [block.id, { ...block, index }]));
    const reads = new Map();
    let sequence = 0;

    function add(observed) {
        for (const item of observed) {
            const original = originals.get(item.id);
            if (item.originalText === false || !original || !Number.isInteger(item.offset) || item.offset < 0
                || !item.text || original.text.slice(item.offset, item.offset + item.text.length) !== item.text) continue;
            const previous = reads.get(item.id) || { ranges: [], lastRead: 0 };
            const ranges = [...previous.ranges, [item.offset, item.offset + item.text.length]].sort((a, b) => a[0] - b[0]);
            const merged = [];
            for (const range of ranges) {
                const last = merged[merged.length - 1];
                if (last && range[0] <= last[1]) last[1] = Math.max(last[1], range[1]);
                else merged.push([...range]);
            }
            reads.set(item.id, { ranges: merged, lastRead: ++sequence });
        }
    }

    function select({ query = '', evidenceIds = [], maxChars = DOCUMENT_EVIDENCE_LIMITS.chars } = {}) {
        if (!Array.isArray(evidenceIds) || evidenceIds.length > DOCUMENT_EVIDENCE_LIMITS.pinnedBlocks
            || evidenceIds.some((id) => typeof id !== 'string' || !originals.has(id))) {
            throw new Error('evidenceIds must contain up to 24 original block IDs.');
        }
        if (!Number.isInteger(maxChars) || maxChars < 2) throw new Error('Invalid review evidence budget.');
        const pinned = new Set(evidenceIds);
        for (const id of pinned) {
            const read = reads.get(id);
            if (!read || read.ranges.length !== 1 || read.ranges[0][0] !== 0
                || read.ranges[0][1] !== originals.get(id).text.length) {
                throw new Error(`Read the complete original block ${id} before pinning it as review evidence.`);
            }
        }
        const candidates = [...reads].map(([id, read]) => {
            const original = originals.get(id);
            const excerpts = read.ranges.map(([start, end]) => ({ id,
                text: original.text.slice(start, end), offset: start,
                nextOffset: end < original.text.length ? end : null }));
            return { id, index: original.index, lastRead: read.lastRead, excerpts,
                words: terms(excerpts.map((item) => item.text).join(' ')), score: 0 };
        });
        const frequencies = new Map();
        for (const candidate of candidates) for (const word of candidate.words) {
            frequencies.set(word, (frequencies.get(word) || 0) + 1);
        }
        const requested = terms(query);
        for (const candidate of candidates) for (const word of requested) {
            if (candidate.words.has(word)) candidate.score += Math.log(1 + candidates.length / frequencies.get(word));
        }
        // Facts relevant to the proposed prose outrank early, unrelated reads.
        // Re-reading a source refreshes its priority without duplicating it.
        candidates.sort((a, b) => Number(pinned.has(b.id)) - Number(pinned.has(a.id))
            || b.score - a.score || b.lastRead - a.lastRead);
        let chars = 2;
        const selected = [];
        const omitted = [];
        for (const candidate of candidates) {
            const cost = JSON.stringify(candidate.excerpts).length - 2
                + (selected.length ? 1 : 0);
            if (chars + cost > maxChars) {
                if (pinned.has(candidate.id)) throw new Error('Pinned review evidence exceeds the budget. Pin fewer blocks or stage a smaller edit.');
                omitted.push(candidate.id);
                continue;
            }
            selected.push(candidate);
            chars += cost;
        }
        selected.sort((a, b) => a.index - b.index);
        return { blocks: selected.flatMap((item) => item.excerpts), coverage: {
            observedBlocks: reads.size, includedBlocks: selected.length, omittedBlocks: omitted.length,
            omittedIds: omitted.slice(0, 24), maxChars,
            note: omitted.length ? 'Some observed blocks were omitted from this review. Do not assume missing evidence is absent from the document. Request relevant original block IDs in neededEvidenceIds.' : '',
        } };
    }
    return { add, select, get size() { return reads.size; } };
}
