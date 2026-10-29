/** Shared native command for explicitly requested or opted-in undo cleanup. */
export function undoHistoryCapability() {
    const context = typeof Office === 'undefined' ? null : Office.context;
    const platform = context?.platform || 'Unknown';
    const version = context?.diagnostics?.version || 'unknown version';
    const supported = context?.requirements?.isSetSupported?.('WordApiDesktop', '1.4') === true;
    const runtimeAvailable = typeof Word !== 'undefined' && typeof Word.run === 'function';
    // Some Mac hosts report false despite meeting the documented minimum:
    // https://github.com/OfficeDev/office-js/issues/6375. This is not proof
    // that undoClear works; only a successful native sync can establish that.
    const canAttempt = runtimeAvailable && (supported || platform === 'Mac');
    return { supported, canAttempt, description: `${platform}, Word ${version}, WordApiDesktop 1.4: ${supported ? 'supported' : 'not reported'}` };
}

export function assertUndoHistoryAvailable() {
    const capability = undoHistoryCapability();
    if (!capability.canAttempt) {
        throw new Error(`Clearing undo history requires a desktop version of Word that supports WordApiDesktop 1.4 (${capability.description}).`);
    }
    return capability;
}

export async function clearUndoHistory(log = () => {}) {
    const capability = assertUndoHistoryAvailable();
    if (!capability.supported) {
        log(`Word's capability report is inconclusive (${capability.description}); attempting the native undoClear command.`, 'info');
    }
    const started = performance.now();
    try {
        await Word.run(async (context) => {
            if (typeof context.document.undoClear !== 'function') {
                throw new Error('This Word runtime does not provide the clear undo history command.');
            }
            context.document.undoClear();
            await context.sync();
        });
    } catch (error) {
        const code = error.code ? ` [${error.code}]` : '';
        throw new Error(`${error.message || String(error)}${code} (${capability.description}).`);
    }
    return { durationMs: Math.round(performance.now() - started) };
}

/** Maintenance failures must never turn a verified edit into a failed edit. */
export async function clearUndoHistoryAfterChangeSet(config, log = () => {}) {
    if (config?.clearUndoHistoryAfterApply !== true) return null;
    let result;
    try {
        const { durationMs } = await clearUndoHistory(log);
        result = { type: 'success', message: `Undo history cleared after this ChangeSet (${durationMs} ms).` };
    } catch (error) {
        result = { type: 'warning', message: `ChangeSet applied, but undo history was not cleared: ${error.message || String(error)}` };
    }
    log(result.message, result.type);
    return result;
}
