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
 * @param {function(): boolean} [deps.getAutoApply] - Current auto-apply setting
 * @param {function(boolean)} [deps.setAutoApply] - Persists an auto-apply change
 * @returns {{ setProcessing: function(boolean), setValue: function(string), focus: function(), updateModelPill: function(string), setSelectionPreview: function(object|string) }}
 */
export function initInputBar({ onSubmit, onCancel, getSkills, onOpenSettings, getAutoApply, setAutoApply }) {
    const textarea = document.getElementById('chatInput');
    const sendBtn = document.getElementById('sendBtn');
    const picker = document.getElementById('skillPicker');
    const modelPill = document.getElementById('modelPill');
    const addSkillBtn = document.getElementById('addSkillBtn');
    const skillsMenu = document.getElementById('skillsMenu');
    const autoApplyToggle = document.getElementById('autoApplyToggle');

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

    // "+" opens a menu of all skills (the "/" picker's click-first twin).
    // Picking one inserts its slash command so the user can add arguments.
    function openSkillsMenu() {
        const skills = getSkills() || [];
        skillsMenu.innerHTML = '';
        if (skills.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'skill-picker-desc';
            empty.textContent = 'No skills available.';
            skillsMenu.appendChild(empty);
        }
        for (const skill of skills) {
            const item = document.createElement('button');
            item.type = 'button';
            item.className = 'skill-picker-item';
            item.setAttribute('role', 'menuitem');

            const name = document.createElement('span');
            name.className = 'skill-picker-name';
            name.textContent = skill.slash;
            const desc = document.createElement('span');
            desc.className = 'skill-picker-desc';
            desc.textContent = skill.description;

            item.appendChild(name);
            item.appendChild(desc);
            item.addEventListener('click', () => {
                closeSkillsMenu();
                pickSkill(skill);
            });
            skillsMenu.appendChild(item);
        }
        skillsMenu.removeAttribute('hidden');
        addSkillBtn.setAttribute('aria-expanded', 'true');
    }

    function closeSkillsMenu() {
        skillsMenu.setAttribute('hidden', '');
        addSkillBtn.setAttribute('aria-expanded', 'false');
    }

    if (addSkillBtn && skillsMenu) {
        addSkillBtn.addEventListener('click', () => {
            if (skillsMenu.hasAttribute('hidden')) {
                closePicker();
                openSkillsMenu();
            } else {
                closeSkillsMenu();
            }
        });
        document.addEventListener('click', (e) => {
            if (!skillsMenu.hasAttribute('hidden') &&
                !skillsMenu.contains(e.target) && e.target !== addSkillBtn) {
                closeSkillsMenu();
            }
        });
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && !skillsMenu.hasAttribute('hidden')) {
                closeSkillsMenu();
            }
        });
    }

    if (autoApplyToggle) {
        if (typeof getAutoApply === 'function') {
            autoApplyToggle.checked = getAutoApply() === true;
        }
        autoApplyToggle.addEventListener('change', () => {
            if (typeof setAutoApply === 'function') setAutoApply(autoApplyToggle.checked);
        });
    }

    modelPill.addEventListener('click', () => onOpenSettings());

    return {
        /** Morphs the send button between Send and Cancel; toggles input disable. */
        setProcessing(isProcessing) {
            processing = isProcessing;
            sendBtn.classList.toggle('cancel-mode', isProcessing);
            sendBtn.textContent = isProcessing ? '■' : '↑';
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
         * Accepts the watcher's content object ({ text, images,
         * totalImages, hasMultiCellTableRegion, tableRegion } — image
         * entries carry dataUrl thumbnails; tableRegion entries carry
         * corner coords R1C1..RnCn) or a plain string (text only).
         * Empty text + no images + no table flag hides it.
         */
        setSelectionPreview(content) {
            const preview = document.getElementById('selectionPreview');
            const previewText = document.getElementById('selectionPreviewText');
            const previewImages = document.getElementById('selectionPreviewImages');
            if (!preview || !previewText || !previewImages) return;
            const text = typeof content === 'string' ? content : (content && content.text) || '';
            const images = content && Array.isArray(content.images) ? content.images : [];
            const total = (content && Number.isInteger(content.totalImages) && content.totalImages > images.length)
                ? content.totalImages : images.length;
            const tableRegion = (content && typeof content.tableRegion === 'object') ? content.tableRegion : null;
            const hasTable = !!(content && content.hasMultiCellTableRegion);
            const trimmed = text.trim();
            if (!trimmed && images.length === 0 && !hasTable) {
                preview.setAttribute('hidden', '');
                return;
            }
            previewText.textContent = trimmed;
            previewText.title = trimmed;
            previewText.toggleAttribute('hidden', !trimmed);
            previewImages.innerHTML = '';
            // Table region badge renders first when present — multi-cell
            // regions route into the table tool session, not the text
            // pipeline, so the badge primes the user to expect ops output.
            if (hasTable) {
                const label = tableRegion
                    ? `Table R${tableRegion.startRow}C${tableRegion.startCol} → R${tableRegion.endRow}C${tableRegion.endCol}`
                    : 'Table region';
                const tag = document.createElement('span');
                tag.className = 'selection-preview-table';
                tag.textContent = label;
                tag.title = label;
                previewImages.appendChild(tag);
            }
            images.forEach((img) => {
                if (!img || !img.dataUrl) return;
                const el = document.createElement('img');
                el.className = 'selection-preview-img';
                el.src = img.dataUrl;
                el.alt = img.altText || 'Selected image';
                el.title = `${img.width || '?'}×${img.height || '?'}pt${img.altText ? ` — ${img.altText}` : ''}`;
                previewImages.appendChild(el);
            });
            if (total > images.length) {
                const more = document.createElement('span');
                more.className = 'selection-preview-more';
                more.textContent = `+${total - images.length}`;
                previewImages.appendChild(more);
            }
            previewImages.toggleAttribute('hidden', previewImages.childNodes.length === 0);
            preview.removeAttribute('hidden');
        },
    };
}
