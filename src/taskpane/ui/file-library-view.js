import { saveFile, listFiles, readFile, getAttachment, renameFile, deleteFile } from '../../lib/file-store.js';
import { formatBytes } from '../../lib/file-attachments.js';
import { containFocus } from './dialog.js';

const PAGE_SIZE = 30;
const PREVIEW_SIZE = 12000;

function element(tag, className, text) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
}

function button(text, action) {
    const node = element('button', 'btn btn-compact', text);
    node.type = 'button';
    node.addEventListener('click', action);
    return node;
}

export function initFileLibraryView({ onAttach } = {}) {
    const overlay = document.getElementById('filesOverlay');
    if (!overlay) return { open: async () => {}, close: () => {} };
    const panel = document.getElementById('filesPanel');
    const search = document.getElementById('filesSearch');
    const kind = document.getElementById('filesKind');
    const upload = document.getElementById('filesUploadBtn');
    const input = document.getElementById('filesInput');
    const list = document.getElementById('filesList');
    const status = document.getElementById('filesStatus');
    const error = document.getElementById('filesError');
    const detail = document.getElementById('filesPreview');
    const previous = document.getElementById('filesPreviousBtn');
    const next = document.getElementById('filesNextBtn');
    let offset = 0;
    let total = 0;
    let request = 0;
    let generation = 0;
    let busy = false;
    let loading = false;
    let releaseFocus;
    let cancelDetail;

    function showError(err) {
        error.textContent = err?.message || String(err || 'File operation failed.');
        error.hidden = false;
    }

    function clearError() {
        error.textContent = '';
        error.hidden = true;
    }

    function controls() {
        upload.disabled = busy;
        search.disabled = busy;
        kind.disabled = busy;
        previous.disabled = busy || loading || offset === 0;
        next.disabled = busy || loading || offset + PAGE_SIZE >= total;
        panel.setAttribute('aria-busy', String(busy || loading));
        list.querySelectorAll('button').forEach((node) => { node.disabled = busy || loading; });
        detail.querySelectorAll('button, input').forEach((node) => { node.disabled = busy; });
    }

    async function refresh() {
        const current = ++request;
        loading = true;
        status.textContent = 'Loading files…';
        controls();
        try {
            const result = await listFiles({ query: search.value.trim(), kind: kind.value || undefined, limit: PAGE_SIZE, offset });
            if (current !== request || overlay.hidden) return;
            total = result.total;
            list.replaceChildren();
            for (const item of result.items) {
                const meta = Object.freeze({ ...item });
                const row = element('li', 'file-library-item');
                const name = button(meta.name, () => preview(meta));
                name.classList.add('file-library-name');
                name.setAttribute('aria-label', `Preview ${meta.name}`);
                row.append(name, element('span', 'help-text', `${String(meta.kind).toUpperCase()} · ${formatBytes(meta.size)}`));
                if (meta.parseError) row.append(element('p', 'input-error', meta.parseError));
                if (meta.persistence === 'memory') row.append(element('p', 'help-text', 'Temporary storage: this file will be lost when this pane reloads.'));
                const actions = element('div', 'file-library-actions');
                actions.append(
                    button('Attach', () => attach(meta)),
                    button('Rename', () => rename(meta)),
                    button('Delete', () => confirmDelete(meta)),
                );
                row.append(actions);
                list.append(row);
            }
            status.textContent = total === 0
                ? 'No files found. Import files to keep them in your library.'
                : `${offset + 1}–${Math.min(offset + result.items.length, total)} of ${total} files`;
        } catch (err) {
            if (current === request && !overlay.hidden) {
                list.replaceChildren();
                status.textContent = 'Could not load files. Search again or reopen Files to retry.';
                showError(err);
            }
        } finally {
            if (current === request) {
                loading = false;
                controls();
            }
        }
    }

    async function run(label, action) {
        if (busy) return;
        const current = generation;
        busy = true;
        clearError();
        status.textContent = label;
        controls();
        try {
            await action(current);
        } catch (err) {
            if (current === generation && !overlay.hidden) {
                showError(err);
                status.textContent = 'File operation failed.';
            }
        } finally {
            busy = false;
            controls();
            if (!overlay.hidden && (!panel.contains(document.activeElement) || document.activeElement.closest('[hidden]'))) search.focus();
        }
    }

    function clearDetail() {
        cancelDetail = null;
        detail.replaceChildren();
        detail.hidden = true;
    }

    function showDetail(meta) {
        clearDetail();
        detail.hidden = false;
        detail.append(element('h3', '', meta.name));
        const close = button('Close preview', () => { clearDetail(); search.focus(); });
        detail.append(close);
        detail.focus();
    }

    async function preview(meta, previewOffset = 0) {
        return run('Loading preview…', async (current) => {
            const result = await readFile(meta.fileId, { offset: previewOffset, limit: PREVIEW_SIZE });
            if (current !== generation || overlay.hidden) return;
            showDetail(result.file || meta);
            if (result.attachment) {
                const dataUrl = result.attachment.dataUrl;
                if (typeof dataUrl !== 'string' || !/^data:image\/(png|jpeg|gif|webp);base64,/i.test(dataUrl)) {
                    throw new Error('This image cannot be previewed safely.');
                }
                const image = element('img', 'file-library-image');
                image.src = dataUrl;
                image.alt = meta.name;
                detail.append(image);
            } else {
                detail.append(element('p', 'help-text', meta.kind === 'text'
                    ? 'Read-only text preview. Your original local file is unchanged.'
                    : 'Extracted text preview, not the original layout. PDF and DOCX files are not edited here.'));
                detail.append(element('pre', 'file-library-text', result.text || '(No text extracted)'));
                const paging = element('div', 'file-library-actions');
                if (previewOffset > 0) paging.append(button('Previous text', () => preview(meta, Math.max(0, previewOffset - PREVIEW_SIZE))));
                if (result.hasMore) paging.append(button('Next text', () => preview(meta, previewOffset + (result.text?.length || PREVIEW_SIZE))));
                detail.append(paging);
            }
            status.textContent = 'Preview ready.';
        });
    }

    async function attach(meta) {
        return run('Preparing attachment…', async (current) => {
            const parsed = await getAttachment(meta.fileId);
            if (current !== generation || overlay.hidden) return;
            if (typeof onAttach !== 'function' || onAttach(parsed) !== true) {
                throw new Error('Could not attach this file. The message may be busy, still loading attachments, or at its attachment limit.');
            }
            close();
            document.getElementById('chatInput')?.focus();
        });
    }

    function rename(meta) {
        if (busy) return;
        showDetail(meta);
        const form = element('form', 'file-library-form');
        const label = element('label', '', 'Library file name');
        const field = element('input', 'form-control');
        field.id = 'filesRenameInput';
        field.value = meta.name;
        field.maxLength = 240;
        field.required = true;
        label.htmlFor = field.id;
        const save = button('Save name', () => {});
        save.type = 'submit';
        form.append(label, field, save);
        form.addEventListener('submit', (event) => {
            event.preventDefault();
            const name = field.value.trim();
            if (!name) { showError(new Error('Enter a file name.')); return; }
            void run('Renaming file…', async (current) => {
                await renameFile(meta.fileId, name);
                if (current !== generation || overlay.hidden) return;
                clearDetail();
                await refresh();
                search.focus();
            });
        });
        detail.append(form, element('p', 'help-text', 'Renames the library entry only. Your original local file is unchanged.'));
        field.focus();
        field.select();
    }

    function confirmDelete(meta) {
        if (busy) return;
        showDetail(meta);
        const cancel = () => { clearDetail(); search.focus(); };
        cancelDetail = cancel;
        detail.append(element('p', '', `Delete “${meta.name}” from the library? This cannot be undone. Your original local file is unchanged.`));
        const actions = element('div', 'file-library-actions');
        const cancelButton = button('Cancel', cancel);
        actions.append(cancelButton, button('Delete file', () => run('Deleting file…', async (current) => {
            await deleteFile(meta.fileId);
            if (current !== generation || overlay.hidden) return;
            clearDetail();
            offset = 0;
            await refresh();
            search.focus();
        })));
        detail.append(actions);
        cancelButton.focus();
    }

    function close() {
        generation += 1;
        request += 1;
        loading = false;
        overlay.hidden = true;
        clearDetail();
        releaseFocus?.();
        releaseFocus = null;
    }

    async function open() {
        if (!overlay.hidden) return;
        generation += 1;
        overlay.hidden = false;
        clearError();
        clearDetail();
        releaseFocus = containFocus(panel, () => {
            if (cancelDetail && !busy) cancelDetail();
            else close();
        }, search);
        offset = 0;
        await refresh();
    }

    document.getElementById('filesCloseBtn').addEventListener('click', close);
    overlay.addEventListener('click', (event) => { if (event.target === overlay) close(); });
    upload.addEventListener('click', () => input.click());
    input.addEventListener('change', () => {
        const files = Array.from(input.files || []);
        input.value = '';
        if (!files.length) return;
        void run('Importing files…', async () => {
            const failures = [];
            for (const file of files) {
                status.textContent = `Importing ${file.name}…`;
                try {
                    const meta = await saveFile(file);
                    if (meta.parseError) failures.push(`${file.name}: saved, but ${meta.parseError}`);
                } catch (err) { failures.push(`${file.name}: ${err?.message || 'Import failed.'}`); }
            }
            if (!overlay.hidden) { offset = 0; await refresh(); }
            if (failures.length) showError(new Error(failures.join('\n')));
        });
    });
    const filter = () => { offset = 0; clearDetail(); clearError(); void refresh(); };
    search.addEventListener('input', filter);
    kind.addEventListener('change', filter);
    previous.addEventListener('click', () => { offset = Math.max(0, offset - PAGE_SIZE); void refresh(); });
    next.addEventListener('click', () => { offset += PAGE_SIZE; void refresh(); });
    return { open, close };
}
