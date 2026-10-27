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
    const selectReport = () => { preview.focus(); preview.select(); preview.setSelectionRange(0, preview.value.length); };
    button.addEventListener('click', async () => {
        button.disabled = true;
        fallback.hidden = true;
        preview.value = '';
        status.textContent = 'Collecting context…';
        try {
            await copyGeneratedText(async () => {
                const report = await getReport();
                preview.value = report;
                return report;
            }, { container: panel });
            status.textContent = 'All context copied.';
            preview.value = '';
        } catch (error) {
            if (preview.value) {
                fallback.hidden = false;
                status.textContent = 'Context is ready. Click Copy prepared text, or press ⌘C / Ctrl+C to copy the selected text.';
                selectReport();
            } else status.textContent = `Could not collect context: ${error.message}`;
        } finally { button.disabled = false; }
    });
    retry.addEventListener('click', async () => {
        try {
            // No document read or module load before this fresh user gesture.
            await copyText(preview.value, { container: panel });
            status.textContent = 'All context copied.';
        } catch {
            status.textContent = 'Press ⌘C / Ctrl+C to copy the selected context.';
            selectReport();
        }
    });
}
