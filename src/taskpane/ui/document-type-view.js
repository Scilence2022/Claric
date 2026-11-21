import { WRITING_OPTIONS } from '../../lib/document-writing-settings.js';

export function initDocumentTypeView(controller) {
    const controls = document.getElementById('chatDocumentTypeControls');
    if (!controls) return;
    const picker = document.getElementById('chatDocumentTypePicker');
    const current = document.getElementById('chatDocumentTypeCurrent');
    const select = document.getElementById('chatDocumentTypeSelect');
    const useDetected = document.getElementById('chatDocumentTypeUseDetected');
    const refresh = document.getElementById('chatDocumentTypeRefresh');
    const detected = document.getElementById('chatDocumentTypeDetected');
    const help = document.getElementById('chatDocumentTypeHelp');

    const labels = WRITING_OPTIONS.type;
    for (const [value, label] of Object.entries(labels)) select.add(new Option(label, value));

    select.addEventListener('change', () => controller.update({
        values: { type: select.value }, locked: { type: Boolean(select.value) },
    }));
    useDetected.addEventListener('click', () => {
        controller.useDetectedType();
        picker.open = false;
        current.focus();
    });
    refresh.addEventListener('click', () => { void controller.detect({ typeOnly: true }); });
    document.addEventListener('pointerdown', (event) => {
        if (picker.open && !controls.contains(event.target)) picker.open = false;
    });
    controls.addEventListener('keydown', (event) => {
        if (event.key === 'Escape' && picker.open) {
            picker.open = false;
            current.focus();
        }
    });

    controller.subscribe(({ settings, detecting, detectionFailed }) => {
        const { detectedType, values, locked } = settings;
        const activeLabel = 'Type: ' + (labels[values.type] || 'Document context');
        current.textContent = current.title = activeLabel + (locked.type ? ' (locked)' : '');
        select.value = values.type;
        useDetected.hidden = !detectedType || (!locked.type && values.type === detectedType);
        useDetected.disabled = detecting;
        refresh.disabled = detecting;
        const label = labels[detectedType] || (detectedType === null ? 'Not checked' : 'No clear type');
        detected.textContent = detecting ? 'Checking type…' : 'Detected: ' + label + (detectionFailed ? ' · Refresh failed' : '');
        help.textContent = locked.type ? 'Locked. Refreshing detection keeps this choice.' : 'Choosing a type locks it.';
    });
    controls.hidden = false;
}
