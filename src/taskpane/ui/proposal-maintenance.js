/** Maintenance status is independent of the edit's verified outcome. */
const statuses = new WeakMap();

export async function notifyChangeSetApplied(el, callback, id) {
    let result;
    try { result = await callback(id); }
    catch (error) { result = { type: 'warning', message: error.message || String(error) }; }
    if (!result?.message) return;
    let state = statuses.get(el);
    if (!state) {
        const status = document.createElement('div');
        status.className = 'proposal-cleanup-status help-text';
        status.setAttribute('role', 'status');
        el.appendChild(status);
        state = { status, successes: 0, failures: 0, lastFailure: '' };
        statuses.set(el, state);
    }
    if (result.type === 'success') state.successes++;
    else { state.failures++; state.lastFailure = result.message; }
    state.status.textContent = state.failures
        ? `Undo cleanup: ${state.successes} succeeded, ${state.failures} failed. ${state.lastFailure}`
        : `Undo history cleared after ${state.successes} ChangeSet(s). ${result.message}`;
}
