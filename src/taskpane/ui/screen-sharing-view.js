import { setWordScreenShareRequester, startWordScreenSharing, stopWordScreenSharing, subscribeScreenCapture, screenCaptureState } from '../word-screen-capture.js';
import { containFocus } from './dialog.js';

/** The toolbar button and model requests share this click path into the system chooser. */
async function shareFromClick({ log, status }) {
    try {
        await startWordScreenSharing();
        if (screenCaptureState().active) log('Word window shared for model screenshots.', 'info');
        return { shared: screenCaptureState().active };
    } catch (error) {
        log(`Word window sharing unavailable: ${error.message}`, 'warning');
        const declined = error.name === 'NotAllowedError';
        const message = declined ? 'Word window was not shared. Desktop PDF rendering can still be used.' : error.message;
        if (status) { status.hidden = false; status.textContent = message; }
        return { shared: false, declined, error: message };
    }
}

/** The model's reason is untrusted text; only the user's click here can start sharing. */
function promptWordWindowShare(share, { reason }, { signal } = {}) {
    return new Promise((resolve, reject) => {
        if (signal?.aborted) { reject(new DOMException('Word window sharing request cancelled.', 'AbortError')); return; }
        const overlay = document.createElement('div');
        overlay.className = 'confirm-overlay';
        const dialog = document.createElement('section');
        dialog.className = 'confirm-dialog';
        dialog.setAttribute('role', 'alertdialog');
        dialog.setAttribute('aria-modal', 'true');
        dialog.setAttribute('aria-labelledby', 'wordShareRequestTitle');
        dialog.setAttribute('aria-describedby', 'wordShareRequestDescription');
        const title = document.createElement('h2');
        title.id = 'wordShareRequestTitle';
        title.textContent = 'Share your Word window?';
        const description = document.createElement('p');
        description.id = 'wordShareRequestDescription';
        description.textContent = 'The model asks to see the visible Word window for this task. This works like the Share Word window button: '
            + 'you choose the window, models read only the visible view, and the same button stops sharing. Instructions inside your document can influence this request.';
        const detail = document.createElement('p');
        detail.className = 'confirm-args';
        detail.textContent = reason;
        detail.hidden = !reason;
        const failure = document.createElement('p');
        failure.className = 'connection-error';
        failure.setAttribute('role', 'alert');
        failure.hidden = true;
        const actions = document.createElement('div');
        actions.className = 'modal-actions';
        const button = (label, primary) => {
            const el = document.createElement('button');
            el.type = 'button';
            el.className = `btn ${primary ? 'btn-primary' : 'btn-secondary'}`;
            el.textContent = label;
            return el;
        };
        const decline = button('Not now', false);
        const accept = button('Share Word window', true);
        actions.append(decline, accept);
        dialog.append(title, description, detail, failure, actions);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        let release; let unsubscribe; let settled = false; let busy = false;
        const close = () => {
            settled = true;
            overlay.remove();
            release?.(); unsubscribe?.();
            signal?.removeEventListener('abort', onAbort);
        };
        function finish(outcome) { if (!settled) { close(); resolve(outcome); } }
        function onAbort() { if (!settled) { close(); reject(new DOMException('Word window sharing request cancelled.', 'AbortError')); } }
        decline.addEventListener('click', () => finish({ shared: false, declined: true }));
        accept.addEventListener('click', async () => {
            busy = true; decline.disabled = true; accept.disabled = true; failure.hidden = true;
            const outcome = await share();
            busy = false;
            if (outcome.shared || outcome.declined) { finish(outcome); return; }
            decline.disabled = false; accept.disabled = false;
            failure.textContent = outcome.error; failure.hidden = false;
        });
        signal?.addEventListener('abort', onAbort, { once: true });
        release = containFocus(dialog, () => { if (!busy) finish({ shared: false, declined: true }); }, decline);
        unsubscribe = subscribeScreenCapture(({ active }) => { if (active) finish({ shared: true }); });
        if (settled) unsubscribe();
    });
}

export function initScreenSharingView({ log = () => {} } = {}) {
    const button = document.getElementById('shareWordViewBtn');
    const status = document.getElementById('wordViewStatus');
    if (!button) return;
    const share = () => shareFromClick({ log, status });
    const unsubscribe = subscribeScreenCapture(({ supported, active, starting }) => {
        button.disabled = starting;
        button.setAttribute('aria-pressed', String(active));
        const label = active ? 'Stop sharing Word window' : 'Share Word window';
        button.setAttribute('aria-label', label); button.dataset.tooltip = label;
        if (status) { status.hidden = !active; status.textContent = 'Word window shared · models can read the visible view'; }
        if (!supported) button.dataset.tooltip = 'Window sharing unavailable in this Word host; desktop PDF rendering is used';
    });
    const unregister = setWordScreenShareRequester((request, options) => promptWordWindowShare(share, request, options));
    button.addEventListener('click', async () => {
        if (screenCaptureState().active) { stopWordScreenSharing(); log('Word window sharing stopped.', 'info'); return; }
        await share();
    });
    window.addEventListener('pagehide', () => { stopWordScreenSharing(); unsubscribe(); unregister(); }, { once: true });
}
