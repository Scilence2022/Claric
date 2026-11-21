import { WRITING_FIELDS, WRITING_OPTIONS } from '../../lib/document-writing-settings.js';

export function initDocumentWritingView(controller) {
    const panel = document.getElementById('settingsPageDocument');
    if (!panel || !controller) return;
    const mode = panel.querySelector('#documentWritingMode');
    const detect = panel.querySelector('#detectDocumentWriting');
    const apply = panel.querySelector('#applyDocumentWriting');
    const undo = panel.querySelector('#undoDocumentWriting');
    const status = panel.querySelector('#documentWritingStatus');
    const suggestion = panel.querySelector('#documentWritingSuggestion');
    const fields = WRITING_FIELDS.map((field) => {
        const select = panel.querySelector(`[data-writing-value="${field}"]`);
        const lock = panel.querySelector(`[data-writing-lock="${field}"]`);
        for (const [value, label] of Object.entries(WRITING_OPTIONS[field])) {
            select.append(new Option(label, value));
        }
        select.addEventListener('change', () => controller.update({
            values: { [field]: select.value }, locked: { [field]: true },
        }));
        lock.addEventListener('change', () => controller.update({ locked: { [field]: lock.checked } }));
        return { field, select, lock };
    });

    mode.addEventListener('change', () => controller.update({ mode: mode.value }));
    detect.addEventListener('click', () => { void controller.detect(); });
    apply.addEventListener('click', () => controller.applySuggestion());
    undo.addEventListener('click', () => controller.undo());

    controller.subscribe(({ settings, detecting, suggestion: proposed, canUndo, status: message }) => {
        mode.value = settings.mode;
        for (const { field, select, lock } of fields) {
            select.value = settings.values[field];
            lock.checked = settings.locked[field];
        }
        detect.disabled = detecting || settings.mode === 'manual' || fields.every(({ lock }) => lock.checked);
        apply.hidden = !proposed;
        undo.hidden = !canUndo;
        suggestion.hidden = !proposed;
        suggestion.replaceChildren(...Object.entries(proposed?.values || {}).map(([field, value]) => {
            const item = document.createElement('li');
            item.textContent = WRITING_OPTIONS[field][value];
            return item;
        }));
        status.textContent = message;
    });
}
