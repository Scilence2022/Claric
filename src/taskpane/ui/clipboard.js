/** Clipboard fallback for Office WebViews without the async Clipboard API. */
export async function copyText(text, { container = document.activeElement?.closest('[role="dialog"], [role="alertdialog"]') || document.body } = {}) {
    const focused = document.activeElement;
    const selection = window.getSelection();
    const ranges = Array.from({ length: selection?.rangeCount || 0 }, (_, i) => selection.getRangeAt(i).cloneRange());
    const field = document.createElement('textarea');
    field.value = text;
    field.setAttribute('aria-label', 'Text to copy');
    field.style.cssText = 'position:fixed;left:-10000px;top:0;';
    // Keep the selection inside the active dialog's focus scope. Appending
    // to body lets its focus trap redirect focus before Word can copy it.
    container.appendChild(field);
    let copied = false;
    try {
        field.focus();
        field.select();
        field.setSelectionRange(0, field.value.length);
        copied = document.execCommand?.('copy') === true;
    } catch { /* Try the async API if the legacy command is unavailable. */
    } finally {
        field.remove();
        focused?.focus();
        if (selection && ranges.length) {
            selection.removeAllRanges();
            ranges.forEach((range) => selection.addRange(range));
        }
    }
    if (copied) return;
    try {
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(text);
            return;
        }
    } catch { /* The caller can offer the prepared text for manual copy. */ }
    throw new Error('Automatic clipboard access is unavailable.');
}

/** Register a promised payload during the click, before async Word reads
 * consume WebKit's transient clipboard activation. Generate it only once. */
export async function copyGeneratedText(generate, options = {}) {
    let pending;
    const text = () => pending ||= Promise.resolve().then(generate);
    if (navigator.clipboard?.write && typeof ClipboardItem !== 'undefined') {
        try {
            const data = text().then((value) => new Blob([value], { type: 'text/plain' }));
            // Some hosts reject the constructor/write before consuming data.
            // The original text promise still reports collection failures below.
            data.catch(() => {});
            const item = new ClipboardItem({ 'text/plain': data });
            await navigator.clipboard.write([item]);
            return;
        } catch { /* Reuse the prepared report for the legacy path. */ }
    }
    await copyText(await text(), options);
}
