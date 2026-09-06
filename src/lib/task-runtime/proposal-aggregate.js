/** Collects task proposals while preserving task identity. */
export function createProposalAggregate(graphId) {
    const proposals = [];
    return {
        add(proposal, task = {}) { proposals.push({ ...proposal, graphId, taskId: task.taskId, attemptId: task.attemptId }); return proposals[proposals.length - 1]; },
        list() { return [...proposals]; },
        get size() { return proposals.length; },
    };
}
