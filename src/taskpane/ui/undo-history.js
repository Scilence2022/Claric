/** Manual document maintenance; never called by the changeset apply path. */
import { appState } from '../app-state.js';
import { containFocus } from './dialog.js';
import { addLog } from './status-bar.js';

function confirmClearUndoHistory() {
    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'confirm-overlay';
        const dialog = document.createElement('section');
        dialog.className = 'confirm-dialog';
        dialog.setAttribute('role', 'alertdialog');
        dialog.setAttribute('aria-modal', 'true');
        dialog.setAttribute('aria-labelledby', 'clearUndoConfirmTitle');
        dialog.setAttribute('aria-describedby', 'clearUndoConfirmDescription');
        dialog.innerHTML = '<h2 id="clearUndoConfirmTitle">Clear undo history?</h2>'
            + '<p id="clearUndoConfirmDescription">This permanently clears all Undo history in the current Word document, including edits made before using Claric. Previous edits can no longer be undone. Document content and existing tracked changes are preserved.</p>'
            + '<div class="modal-actions"><button type="button" class="btn btn-secondary" data-confirm="cancel">Cancel</button>'
            + '<button type="button" class="btn btn-primary" data-confirm="clear-undo">Clear undo history</button></div>';
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        let release;
        const finish = (accepted) => {
            overlay.remove();
            release?.();
            resolve(accepted);
        };
        const cancel = dialog.querySelector('[data-confirm="cancel"]');
        cancel.addEventListener('click', () => finish(false));
        dialog.querySelector('[data-confirm="clear-undo"]').addEventListener('click', () => finish(true));
        release = containFocus(dialog, () => finish(false), cancel);
    });
}

function assertAvailable() {
    if (appState.isProcessing || appState.isProcessingDoc || appState.isProcessingSummary
        || appState.processDocController || appState.chatController) {
        throw new Error('Wait for Claric to finish processing or applying changes before clearing undo history.');
    }
    if (typeof Office === 'undefined' || !Office.context?.requirements?.isSetSupported('WordApiDesktop', '1.4')
        || typeof Word === 'undefined' || typeof Word.run !== 'function') {
        throw new Error('Clearing undo history requires a desktop version of Word that supports WordApiDesktop 1.4.');
    }
}

/** Wire the manual Settings button without adding a persistent preference. */
export function initUndoHistoryControl() {
    const button = document.getElementById('clearUndoHistoryBtn');
    const status = document.getElementById('clearUndoHistoryStatus');
    if (!button || !status) return;
    let pending = false;
    const report = (message, type) => {
        status.textContent = message;
        status.hidden = false;
        addLog(message, type);
    };
    button.addEventListener('click', async () => {
        if (pending) return;
        pending = true;
        let ownsBusy = false;
        try {
            assertAvailable();
            if (!await confirmClearUndoHistory()) return;
            // Processing may have started while the confirmation was open.
            assertAvailable();
            appState.isProcessing = true;
            ownsBusy = true;
            button.disabled = true;
            button.setAttribute('aria-busy', 'true');
            status.textContent = 'Clearing undo history...';
            status.hidden = false;
            const started = performance.now();
            await Word.run(async (context) => {
                if (typeof context.document.undoClear !== 'function') {
                    throw new Error('This Word version does not provide the clear undo history command.');
                }
                context.document.undoClear();
                await context.sync();
            });
            report(`Undo history cleared (${Math.round(performance.now() - started)} ms). Existing tracked changes are preserved. Test typing and scrolling in Word to compare responsiveness.`, 'success');
        } catch (error) {
            report(`Could not clear undo history: ${error.message || String(error)}`, 'error');
        } finally {
            if (ownsBusy) appState.isProcessing = false;
            button.disabled = false;
            button.removeAttribute('aria-busy');
            pending = false;
        }
    });
}
