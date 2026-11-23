/** Resolve a contiguous block move against the order at this step of the draft. */
export function resolveDocumentMove(order, { blockIds, targetId, side }) {
    if (!Array.isArray(blockIds) || !blockIds.length || !['before', 'after'].includes(side)) {
        throw new Error('A move needs a contiguous source and exactly one destination side.');
    }
    const start = order.indexOf(blockIds[0]);
    const target = order.indexOf(targetId);
    const end = start + blockIds.length;
    if (start < 0 || blockIds.some((id, index) => order[start + index] !== id)) {
        throw new Error('Move source paragraphs must be contiguous and in their current order.');
    }
    if (target < 0 || (target >= start && target < end)) throw new Error('Move destination must be outside the source.');
    const gap = target + (side === 'after' ? 1 : 0);
    if (gap === start || gap === end) throw new Error('The source already occupies this position.');
    return { start, end, gap, destination: gap > end ? gap - blockIds.length : gap };
}

export function applyDocumentMove(order, move) {
    const { start, destination } = resolveDocumentMove(order, move);
    const next = order.slice();
    next.splice(start, move.blockIds.length);
    next.splice(destination, 0, ...move.blockIds);
    return next;
}
