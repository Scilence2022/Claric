/** Clipboard fallback for Office WebViews without the async Clipboard API. */
export async function copyText(text) {
    try {
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(text);
            return;
        }
    } catch { /* Try the host's legacy copy command. */ }
    const focused = document.activeElement;
    const selection = window.getSelection();
    const ranges = Array.from({ length: selection?.rangeCount || 0 }, (_, i) => selection.getRangeAt(i).cloneRange());
    const field = document.createElement('textarea');
    field.value = text;
    field.setAttribute('aria-label', 'Text to copy');
    field.style.cssText = 'position:fixed;left:-10000px;top:0;';
    document.body.appendChild(field);
    try {
        field.select();
        if (!document.execCommand?.('copy')) throw new Error('Clipboard unavailable. Please try copying again.');
    } finally {
        field.remove();
        focused?.focus();
        if (selection && ranges.length) {
            selection.removeAllRanges();
            ranges.forEach((range) => selection.addRange(range));
        }
    }
}
