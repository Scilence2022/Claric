import { startWordScreenSharing, stopWordScreenSharing, subscribeScreenCapture, screenCaptureState } from '../word-screen-capture.js';

export function initScreenSharingView({ log = () => {} } = {}) {
    const button = document.getElementById('shareWordViewBtn');
    const status = document.getElementById('wordViewStatus');
    if (!button) return;
    const unsubscribe = subscribeScreenCapture(({ supported, active, starting }) => {
        button.disabled = starting;
        button.setAttribute('aria-pressed', String(active));
        const label = active ? 'Stop sharing Word window' : 'Share Word window';
        button.setAttribute('aria-label', label); button.dataset.tooltip = label;
        if (status) { status.hidden = !active; status.textContent = 'Word window shared · models can read the visible view'; }
        if (!supported) button.dataset.tooltip = 'Window sharing unavailable in this Word host; desktop PDF rendering is used';
    });
    button.addEventListener('click', async () => {
        if (screenCaptureState().active) { stopWordScreenSharing(); log('Word window sharing stopped.', 'info'); return; }
        try { await startWordScreenSharing(); if (screenCaptureState().active) log('Word window shared for model screenshots.', 'info'); }
        catch (error) {
            log(`Word window sharing unavailable: ${error.message}`, 'warning');
            if (status) { status.hidden = false; status.textContent = error.name === 'NotAllowedError'
                ? 'Word window was not shared. Desktop PDF rendering can still be used.' : error.message; }
        }
    });
    window.addEventListener('pagehide', () => { stopWordScreenSharing(); unsubscribe(); }, { once: true });
}
