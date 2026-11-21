import { copyGeneratedText, copyText } from './clipboard.js';

/** Keep clipboard registration synchronous with the user's original click. */
export function initContextCopy({ getReport }) {
    const button = document.getElementById('copyContextBtn');
    const panel = document.getElementById('aboutPanel');
    const status = document.getElementById('contextCopyStatus');
    const fallback = document.getElementById('contextCopyFallback');
    const preview = document.getElementById('contextCopyText');
    const retry = document.getElementById('retryContextCopyBtn');
    if (!button || !preview || !retry) return;
    let generation = 0;
    const selectReport = () => { preview.focus(); preview.select(); preview.setSelectionRange(0, preview.value.length); };
    button.addEventListener('click', async () => {
        const startedIn = generation;
        const isCurrent = () => startedIn === generation;
        button.disabled = true;
        fallback.hidden = true;
        preview.value = '';
        status.textContent = 'Collecting context…';
        try {
            await copyGeneratedText(async () => {
                if (!isCurrent()) throw new Error('The conversation changed.');
                const report = await getReport();
                if (!isCurrent()) throw new Error('The conversation changed.');
                preview.value = report;
                return report;
            }, { container: panel, isCurrent });
            if (!isCurrent()) return;
            status.textContent = 'All context copied.';
            preview.value = '';
        } catch (error) {
            if (!isCurrent()) return;
            if (preview.value) {
                fallback.hidden = false;
                status.textContent = 'Context is ready. Click Copy prepared text, or press ⌘C / Ctrl+C to copy the selected text.';
                selectReport();
            } else status.textContent = `Could not collect context: ${error.message}`;
        } finally { if (isCurrent()) button.disabled = false; }
    });
    retry.addEventListener('click', async () => {
        if (!preview.value) return;
        const startedIn = generation;
        try {
            // No document read or module load before this fresh user gesture.
            await copyText(preview.value, { container: panel });
            if (startedIn !== generation) return;
            status.textContent = 'All context copied.';
            preview.value = '';
            fallback.hidden = true;
        } catch {
            if (startedIn !== generation) return;
            status.textContent = 'Press ⌘C / Ctrl+C to copy the selected context.';
            selectReport();
        }
    });
    return () => {
        generation++;
        preview.value = '';
        fallback.hidden = true;
        status.textContent = '';
        button.disabled = false;
    };
}
