/**
 * Input Bar Module
 *
 * Bottom chat input: auto-growing textarea, slash-command skill picker popup,
 * send button that morphs into Cancel while a turn is processing, and the
 * model pill (click opens Settings).
 *
 * @module ui/input-bar
 */

/**
 * Initializes the input bar.
 *
 * @param {object} deps
 * @param {function(string)} deps.onSubmit - Called with the raw input text
 * @param {function()} deps.onCancel - Called when the morphed Cancel button is clicked
 * @param {function(): Array<object>} deps.getSkills - Returns the current skill list
 * @param {function()} deps.onOpenSettings - Opens the settings slide-over
 * @returns {{ setProcessing: function(boolean), setValue: function(string), focus: function(), updateModelPill: function(string), setSelectionPreview: function(string) }}
 */
export function initInputBar({ onSubmit, onCancel, getSkills, onOpenSettings }) {
    const textarea = document.getElementById('chatInput');
    const sendBtn = document.getElementById('sendBtn');
    const picker = document.getElementById('skillPicker');
    const modelPill = document.getElementById('modelPill');

    let processing = false;
    let pickerItems = [];
    let pickerIndex = 0;

    /** Filters skills by the text after '/' and re-renders the picker. */
    function refreshPicker(filter) {
        const skills = (getSkills() || []).filter((s) =>
            s.slash.toLowerCase().includes(filter.toLowerCase()) ||
            s.name.toLowerCase().includes(filter.toLowerCase())
        );
        pickerItems = skills;
        pickerIndex = 0;
        picker.innerHTML = '';
        if (skills.length === 0) {
            picker.setAttribute('hidden', '');
            return;
        }
        skills.forEach((skill, i) => {
            const item = document.createElement('button');
            item.type = 'button';
            item.className = 'skill-picker-item' + (i === pickerIndex ? ' active' : '');

            const name = document.createElement('span');
            name.className = 'skill-picker-name';
            name.textContent = skill.slash;
            const desc = document.createElement('span');
            desc.className = 'skill-picker-desc';
            desc.textContent = skill.description;

            item.appendChild(name);
            item.appendChild(desc);
            item.addEventListener('click', () => pickSkill(skill));
            picker.appendChild(item);
        });
        picker.removeAttribute('hidden');
    }

    /** Inserts the chosen skill's slash command into the textarea. */
    function pickSkill(skill) {
        textarea.value = `${skill.slash} `;
        closePicker();
        textarea.focus();
        autosize();
    }

    function closePicker() {
        picker.setAttribute('hidden', '');
        pickerItems = [];
    }

    /** Moves the picker highlight by delta and updates item classes. */
    function movePicker(delta) {
        if (pickerItems.length === 0) return;
        pickerIndex = (pickerIndex + delta + pickerItems.length) % pickerItems.length;
        picker.querySelectorAll('.skill-picker-item').forEach((el, i) => {
            el.classList.toggle('active', i === pickerIndex);
        });
    }

    /** Grows the textarea up to a cap. */
    function autosize() {
        textarea.style.height = 'auto';
        textarea.style.height = `${Math.min(textarea.scrollHeight, 140)}px`;
    }

    function submitCurrent() {
        const text = textarea.value;
        if (!text.trim() || processing) return;
        closePicker();
        onSubmit(text);
    }

    textarea.addEventListener('input', () => {
        autosize();
        const value = textarea.value;
        if (value.startsWith('/')) {
            const spaceIdx = value.indexOf(' ');
            const filter = spaceIdx === -1 ? value.slice(1) : value.slice(1, spaceIdx);
            refreshPicker(filter);
        } else {
            closePicker();
        }
    });

    textarea.addEventListener('keydown', (e) => {
        const pickerOpen = !picker.hasAttribute('hidden');
        if (pickerOpen) {
            if (e.key === 'ArrowDown') { e.preventDefault(); movePicker(1); return; }
            if (e.key === 'ArrowUp') { e.preventDefault(); movePicker(-1); return; }
            if (e.key === 'Tab' || (e.key === 'Enter' && pickerItems.length > 0)) {
                e.preventDefault();
                pickSkill(pickerItems[pickerIndex]);
                return;
            }
            if (e.key === 'Escape') { closePicker(); return; }
        }
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            submitCurrent();
        }
    });

    sendBtn.addEventListener('click', () => {
        if (processing) {
            onCancel();
        } else {
            submitCurrent();
        }
    });

    // Clicking anywhere else closes the picker.
    document.addEventListener('click', (e) => {
        if (!picker.hasAttribute('hidden') && !picker.contains(e.target) && e.target !== textarea) {
            closePicker();
        }
    });

    modelPill.addEventListener('click', () => onOpenSettings());

    return {
        /** Morphs the send button between Send and Cancel; toggles input disable. */
        setProcessing(isProcessing) {
            processing = isProcessing;
            sendBtn.classList.toggle('cancel-mode', isProcessing);
            sendBtn.textContent = isProcessing ? '✕' : '↑';
            sendBtn.title = isProcessing ? 'Cancel' : 'Send';
            textarea.disabled = isProcessing;
        },
        /** Sets the textarea content (used by skill chips). */
        setValue(text) {
            textarea.value = text;
            autosize();
            closePicker();
        },
        /** Focuses the textarea. */
        focus() {
            textarea.focus();
        },
        /** Updates the model pill label ("Provider: model"). */
        updateModelPill(text) {
            modelPill.textContent = text;
        },
        /**
         * Shows/hides the live selection preview above the input.
         * Pass the current selection text, or ''/null to hide.
         */
        setSelectionPreview(text) {
            const preview = document.getElementById('selectionPreview');
            const previewText = document.getElementById('selectionPreviewText');
            if (!preview || !previewText) return;
            const trimmed = (text || '').trim();
            if (!trimmed) {
                preview.setAttribute('hidden', '');
                return;
            }
            previewText.textContent = trimmed;
            previewText.title = trimmed;
            preview.removeAttribute('hidden');
        },
    };
}
