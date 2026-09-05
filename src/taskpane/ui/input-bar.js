/**
 * Input Bar Module
 *
 * Bottom chat input: auto-growing textarea, slash-command skill picker popup,
 * send button that morphs into Cancel while a turn is processing, and the
 * model pill (click opens Settings).
 *
 * @module ui/input-bar
 */

import {
    ATTACHMENT_KIND,
    detectAttachmentKind,
    validateAttachment,
    parseAttachment,
    formatBytes,
} from '../../lib/file-attachments.js';
import { confirmAutoApply } from './dialog.js';

/** Type icons for attachment chips (text glyphs, matching the composer style). */
const ATTACHMENT_ICONS = Object.freeze({
    [ATTACHMENT_KIND.TEXT]: '¶',
    [ATTACHMENT_KIND.DOCX]: 'W',
    [ATTACHMENT_KIND.PDF]: '§',
});

/**
 * Initializes the input bar.
 *
 * @param {object} deps
 * @param {function(string, Array<object>)} deps.onSubmit - Called with the raw
 *   input text and the parsed attachments ({name, kind, size, text?, dataUrl?})
 * @param {function()} deps.onCancel - Called when the morphed Cancel button is clicked
 * @param {function(): Array<object>} deps.getSkills - Returns the current skill list
 * @param {function()} deps.onOpenSettings - Opens the settings slide-over
 * @param {function(): boolean} [deps.getAutoApply] - Current auto-apply setting
 * @param {function(): boolean} [deps.getTrackChanges] - Current tracking setting
 * @param {function(boolean)} [deps.setAutoApply] - Persists an auto-apply change
 * @param {function(string, string)} [deps.onLog] - Activity-log sink for
 *   attachment validation/parse failures
 * @returns {{ setProcessing: function(boolean), setValue: function(string), focus: function(), setSelectionPreview: function(object|string), clearAttachments: function() }}
 */
export function initInputBar({ onSubmit, onCancel, getSkills, onOpenSettings, getAutoApply, getTrackChanges, setAutoApply, onLog }) {
    const textarea = document.getElementById('chatInput');
    const sendBtn = document.getElementById('sendBtn');
    const picker = document.getElementById('skillPicker');
    const modelPill = document.getElementById('modelPill');
    const addSkillBtn = document.getElementById('addSkillBtn');
    const skillsMenu = document.getElementById('skillsMenu');
    const autoApplyToggle = document.getElementById('autoApplyToggle');
    const attachBtn = document.getElementById('attachBtn');
    const attachInput = document.getElementById('attachmentInput');
    const chipsEl = document.getElementById('attachmentChips');
    const errorEl = document.getElementById('inputError');

    function showError(message) {
        if (!errorEl) return;
        errorEl.textContent = String(message || 'Something went wrong.');
        errorEl.hidden = false;
    }

    function clearError() {
        if (!errorEl) return;
        errorEl.textContent = '';
        errorEl.hidden = true;
    }

    // Parsed attachments pending submission ({name, kind, size, text?,
    // dataUrl?}). Cleared on submit and on new chat.
    let attachments = [];

    let processing = false;
    let composing = false;
    let pickerItems = [];
    let pickerIndex = 0;

    // Submitted-prompt history for the ↑/↓ recall (terminal-style).
    // In-memory per taskpane session; duplicates collapse, draft preserved.
    const inputHistory = [];
    const MAX_INPUT_HISTORY = 100;
    let historyIndex = null;
    let historyDraft = '';

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
            item.id = `skill-picker-option-${i}`;
            item.setAttribute('role', 'option');
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
        picker.setAttribute('aria-activedescendant', `skill-picker-option-${pickerIndex}`);
        textarea.setAttribute('aria-controls', 'skillPicker');
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
        picker.removeAttribute('aria-activedescendant');
        textarea.removeAttribute('aria-controls');
        pickerItems = [];
    }

    /** Moves the picker highlight by delta and updates item classes. */
    function movePicker(delta) {
        if (pickerItems.length === 0) return;
        pickerIndex = (pickerIndex + delta + pickerItems.length) % pickerItems.length;
        picker.setAttribute('aria-activedescendant', `skill-picker-option-${pickerIndex}`);
        picker.querySelectorAll('.skill-picker-item').forEach((el, i) => {
            const active = i === pickerIndex;
            el.classList.toggle('active', active);
            // Keep the keyboard-highlighted item visible when the list
            // overflows the picker's max-height and scrolls.
            if (active && typeof el.scrollIntoView === 'function') {
                el.scrollIntoView({ block: 'nearest' });
            }
        });
    }

    /** Grows the textarea up to a cap. */
    function autosize() {
        textarea.style.height = 'auto';
        textarea.style.height = `${Math.min(textarea.scrollHeight, 140)}px`;
    }

    function recordHistory(text) {
        if (inputHistory[inputHistory.length - 1] !== text) {
            inputHistory.push(text);
            if (inputHistory.length > MAX_INPUT_HISTORY) inputHistory.shift();
        }
        historyIndex = null;
        historyDraft = '';
    }

    function submitCurrent() {
        const text = textarea.value;
        if (parsing) {
            showError([...attachmentErrors, 'Attachments are still loading. Wait before sending.'].join('\n'));
            return;
        }
        if ((!text.trim() && attachments.length === 0) || processing) return;
        closePicker();
        recordHistory(text);
        const sent = attachments;
        attachments = [];
        renderChips();
        onSubmit(text, sent);
    }

    /** Re-renders the attachment chip list above the composer. */
    function renderChips() {
        if (!chipsEl) return;
        chipsEl.innerHTML = '';
        if (attachments.length === 0) {
            chipsEl.setAttribute('hidden', '');
            return;
        }
        attachments.forEach((att, index) => {
            const chip = document.createElement('span');
            chip.className = 'attachment-chip';
            chip.title = `${att.name} (${formatBytes(att.size)})`;

            if (att.kind === ATTACHMENT_KIND.IMAGE && att.dataUrl) {
                const thumb = document.createElement('img');
                thumb.className = 'attachment-chip-thumb';
                thumb.src = att.dataUrl;
                thumb.alt = att.name;
                chip.appendChild(thumb);
            } else {
                const icon = document.createElement('span');
                icon.className = 'attachment-chip-icon';
                icon.setAttribute('aria-hidden', 'true');
                icon.textContent = ATTACHMENT_ICONS[att.kind] || '¶';
                chip.appendChild(icon);
            }

            const name = document.createElement('span');
            name.className = 'attachment-chip-name';
            name.textContent = att.name;
            chip.appendChild(name);

            const remove = document.createElement('button');
            remove.type = 'button';
            remove.className = 'attachment-chip-remove';
            remove.textContent = '×';
            remove.title = 'Remove attachment';
            remove.setAttribute('aria-label', `Remove ${att.name}`);
            remove.addEventListener('click', () => {
                attachments.splice(index, 1);
                renderChips();
            });
            chip.appendChild(remove);
            chipsEl.appendChild(chip);
        });
        chipsEl.removeAttribute('hidden');
    }

    let attachmentGeneration = 0;
    let parsing = false;
    let attachmentErrors = [];

    /** Validates and parses each picked file; failures remain visible until reset. */
    async function addFiles(fileList) {
        if (parsing || processing) return;
        const generation = attachmentGeneration;
        parsing = true;
        if (attachBtn) attachBtn.disabled = true;
        sendBtn.setAttribute('aria-busy', 'true');
        for (const file of Array.from(fileList || [])) {
            if (generation !== attachmentGeneration) return;
            const kind = detectAttachmentKind(file.name, file.type);
            const verdict = validateAttachment({ name: file.name, size: file.size, kind }, attachments);
            if (!verdict.ok) {
                attachmentErrors.push(verdict.error);
                showError(attachmentErrors.join('\n'));
                if (typeof onLog === 'function') onLog(verdict.error, 'warning');
                continue;
            }
            try {
                const parsed = await parseAttachment(file);
                if (generation !== attachmentGeneration) return;
                attachments.push(parsed);
            } catch (err) {
                if (generation !== attachmentGeneration) return;
                attachmentErrors.push(err.message);
                showError(attachmentErrors.join('\n'));
                if (typeof onLog === 'function') onLog(err.message, 'error');
            }
        }
        parsing = false;
        if (attachBtn) attachBtn.disabled = processing;
        sendBtn.removeAttribute('aria-busy');
        if (attachmentErrors.length) showError(attachmentErrors.join('\n'));
        else clearError();
        renderChips();
    }

    if (attachBtn && attachInput) {
        attachBtn.addEventListener('click', () => attachInput.click());
        attachInput.addEventListener('change', () => {
            addFiles(attachInput.files);
            attachInput.value = ''; // re-picking the same file must re-fire change
        });
    }

    textarea.addEventListener('compositionstart', () => {
        composing = true;
        closePicker();
    });

    textarea.addEventListener('compositionend', () => {
        composing = false;
        textarea.dispatchEvent(new Event('input', { bubbles: true }));
    });

    textarea.addEventListener('input', () => {
        autosize();
        historyIndex = null; // manual edit — ↑ restarts from the newest entry
        if (composing) return;
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
        if (composing || e.isComposing || e.keyCode === 229) return;
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
        // ↑/↓ recall submitted prompts (only when the picker is closed and
        // the caret sits at the very start/end, so multi-line caret moves
        // keep working).
        const caretOnFirstLine = textarea.value.slice(0, textarea.selectionStart).indexOf('\n') === -1;
        const caretOnLastLine = textarea.value.slice(textarea.selectionEnd).indexOf('\n') === -1;
        if (e.key === 'ArrowUp' && inputHistory.length > 0 && caretOnFirstLine) {
            e.preventDefault();
            if (historyIndex === null) {
                historyDraft = textarea.value;
                historyIndex = inputHistory.length - 1;
            } else if (historyIndex > 0) {
                historyIndex -= 1;
            }
            textarea.value = inputHistory[historyIndex];
            textarea.setSelectionRange(0, 0);
            autosize();
            return;
        }
        if (e.key === 'ArrowDown' && historyIndex !== null && caretOnLastLine) {
            e.preventDefault();
            historyIndex += 1;
            if (historyIndex >= inputHistory.length) {
                historyIndex = null;
                textarea.value = historyDraft;
            } else {
                textarea.value = inputHistory[historyIndex];
                textarea.setSelectionRange(textarea.value.length, textarea.value.length);
            }
            autosize();
            return;
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
            empty.textContent = 'No slash commands available.';
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
        let confirming = false;
        autoApplyToggle.addEventListener('change', async () => {
            if (confirming) return;
            if (autoApplyToggle.checked) {
                autoApplyToggle.checked = false;
                if (getTrackChanges?.() === false) {
                    showError('Auto-apply requires Track Changes. Enable Track Changes in Settings first.');
                    return;
                }
                confirming = true;
                let confirmed = false;
                try {
                    confirmed = await confirmAutoApply();
                } catch (_error) {
                    showError('Confirmation unavailable. Auto-apply remains disabled.');
                } finally {
                    confirming = false;
                }
                if (!confirmed) return;
                if (getTrackChanges?.() === false) {
                    showError('Auto-apply requires Track Changes. Enable Track Changes in Settings first.');
                    return;
                }
                autoApplyToggle.checked = true;
            }
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
            sendBtn.setAttribute('aria-label', isProcessing ? 'Cancel' : 'Send');
            textarea.disabled = isProcessing;
            if (attachBtn) attachBtn.disabled = isProcessing || parsing;
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
        /** Drops pending attachments (new chat / history switch). */
        clearAttachments() {
            attachmentGeneration += 1;
            parsing = false;
            attachmentErrors = [];
            clearError();
            sendBtn.removeAttribute('aria-busy');
            if (attachBtn) attachBtn.disabled = processing;
            attachments = [];
            renderChips();
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
