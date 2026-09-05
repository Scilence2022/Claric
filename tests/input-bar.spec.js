/** @jest-environment jsdom */

/**
 * Specs for the selection preview chip (ui/input-bar.js): text snippet,
 * image thumbnails, mixed content, +N badge, and hide-on-empty.
 */

const { initInputBar } = require('../src/taskpane/ui/input-bar.js');

function setupDom() {
    document.body.innerHTML = `
        <textarea id="chatInput"></textarea>
        <button id="sendBtn"></button>
        <div id="skillPicker" hidden></div>
        <div id="skillsMenu" hidden></div>
        <button id="addSkillBtn"></button>
        <input type="checkbox" id="autoApplyToggle">
        <button id="modelPill"></button>
        <div id="selectionPreview" hidden>
            <span class="selection-preview-icon"></span>
            <span id="selectionPreviewText"></span>
            <span id="selectionPreviewImages" hidden></span>
        </div>`;
    return initInputBar({ onSubmit: jest.fn(), onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
}

describe('setSelectionPreview', () => {
    test('text-only selection shows the snippet, no thumbnails', () => {
        const bar = setupDom();
        bar.setSelectionPreview({ text: 'some selected words', images: [], totalImages: 0 });

        const preview = document.getElementById('selectionPreview');
        expect(preview.hasAttribute('hidden')).toBe(false);
        expect(document.getElementById('selectionPreviewText').textContent).toBe('some selected words');
        expect(document.getElementById('selectionPreviewImages').hasAttribute('hidden')).toBe(true);
    });

    test('image-only selection shows thumbnails without text', () => {
        const bar = setupDom();
        bar.setSelectionPreview({
            text: '',
            images: [{ dataUrl: 'data:image/png;base64,iVBOR', width: 320, height: 240, altText: 'chart' }],
            totalImages: 1,
        });

        expect(document.getElementById('selectionPreview').hasAttribute('hidden')).toBe(false);
        expect(document.getElementById('selectionPreviewText').hasAttribute('hidden')).toBe(true);
        const imgs = document.querySelectorAll('.selection-preview-img');
        expect(imgs).toHaveLength(1);
        expect(imgs[0].src).toBe('data:image/png;base64,iVBOR');
        expect(imgs[0].alt).toBe('chart');
        expect(imgs[0].title).toContain('320×240pt');
    });

    test('mixed selection shows both; truncated totals render a +N badge', () => {
        const bar = setupDom();
        bar.setSelectionPreview({
            text: 'caption text',
            images: [{ dataUrl: 'data:image/png;base64,iVBOR' }, { dataUrl: 'data:image/jpeg;base64,/9j/' }],
            totalImages: 5,
        });

        expect(document.getElementById('selectionPreviewText').textContent).toBe('caption text');
        expect(document.querySelectorAll('.selection-preview-img')).toHaveLength(2);
        expect(document.querySelector('.selection-preview-more').textContent).toBe('+3');
    });

    test('multi-cell table selection renders corner-coords badge', () => {
        const bar = setupDom();
        bar.setSelectionPreview({
            text: 'R1C1 R1C2 R2C1',
            hasMultiCellTableRegion: true,
            tableRegion: { startRow: 1, endRow: 2, startCol: 1, endCol: 2 },
        });

        expect(document.getElementById('selectionPreview').hasAttribute('hidden')).toBe(false);
        expect(document.querySelector('.selection-preview-table').textContent)
            .toBe('Table R1C1 → R2C2');
    });

    test('table-region flag without coords falls back to a generic label', () => {
        const bar = setupDom();
        bar.setSelectionPreview({
            text: '',
            hasMultiCellTableRegion: true,
            tableRegion: null,
        });

        expect(document.querySelector('.selection-preview-table').textContent).toBe('Table region');
    });

    test('empty content and legacy strings still work', () => {
        const bar = setupDom();
        bar.setSelectionPreview({ text: '', images: [], totalImages: 0 });
        expect(document.getElementById('selectionPreview').hasAttribute('hidden')).toBe(true);

        // Legacy string form (text-only) renders as before.
        bar.setSelectionPreview('plain string');
        expect(document.getElementById('selectionPreviewText').textContent).toBe('plain string');
    });
});


describe('skills "+" menu and auto-apply toggle', () => {
    beforeEach(() => {
        document.body.innerHTML = '<textarea id="chatInput"></textarea><button id="sendBtn"></button><div id="skillPicker" hidden></div><button id="modelPill"></button><input type="checkbox" id="autoApplyToggle"><div id="inputError" hidden></div>';
    });

    test('a failed DOM confirmation never authorizes auto-apply', async () => {
        const setAutoApply = jest.fn();
        initInputBar({ onSubmit: jest.fn(), onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn(), setAutoApply });
        const append = jest.spyOn(document.body, 'appendChild').mockImplementationOnce(() => { throw new Error('Unavailable'); });
        const toggle = document.getElementById('autoApplyToggle');
        toggle.checked = true;
        toggle.dispatchEvent(new Event('change'));
        await Promise.resolve();
        append.mockRestore();
        expect(toggle.checked).toBe(false);
        expect(setAutoApply).not.toHaveBeenCalled();
        expect(document.getElementById('inputError').textContent).toContain('Confirmation unavailable');
    });

    test('tracking disabled refuses auto-apply with an inline reason', () => {
        const setAutoApply = jest.fn();
        initInputBar({ onSubmit: jest.fn(), onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn(), getTrackChanges: () => false, setAutoApply });
        const toggle = document.getElementById('autoApplyToggle');
        toggle.checked = true;
        toggle.dispatchEvent(new Event('change'));
        expect(toggle.checked).toBe(false);
        expect(setAutoApply).not.toHaveBeenCalled();
        expect(document.getElementById('inputError').textContent).toContain('requires Track Changes');
        expect(document.querySelector('[role="alertdialog"]')).toBeNull();
    });
    test('"+" opens a menu of all skills; picking one inserts the slash command', () => {
        document.body.innerHTML = `
            <textarea id="chatInput"></textarea>
            <button id="sendBtn"></button>
            <div id="skillPicker" hidden></div>
            <div id="skillsMenu" hidden></div>
            <button id="addSkillBtn"></button>
            <input type="checkbox" id="autoApplyToggle">
            <button id="modelPill"></button>`;
        const skills = [
            { name: 'mcp', slash: '/mcp', description: 'Run with MCP tools' },
            { name: 'copy-edit', slash: '/copy-edit', description: 'Fix errors' },
        ];
        initInputBar({
            onSubmit: jest.fn(), onCancel: jest.fn(), onOpenSettings: jest.fn(),
            getSkills: () => skills,
        });

        document.getElementById('addSkillBtn').click();
        const menu = document.getElementById('skillsMenu');
        expect(menu.hasAttribute('hidden')).toBe(false);
        const items = menu.querySelectorAll('.skill-picker-item');
        expect(items).toHaveLength(2);
        expect(items[0].textContent).toContain('/mcp');

        const textarea = document.getElementById('chatInput');
        textarea.value = '/m';
        textarea.dispatchEvent(new Event('input', { bubbles: true }));
        const picker = document.getElementById('skillPicker');
        expect(textarea.getAttribute('aria-controls')).toBe('skillPicker');
        expect(picker.getAttribute('aria-activedescendant')).toBe('skill-picker-option-0');
        expect(picker.querySelector('[role="option"]')).not.toBeNull();

        items[0].click();
        expect(menu.hasAttribute('hidden')).toBe(true);
        expect(document.getElementById('chatInput').value).toBe('/mcp ');
    });

    test('auto-apply toggle initializes from config and persists changes', async () => {
        window.confirm = jest.fn(() => { throw new Error('WebView'); });
        let stored = false;
        initInputBar({
            onSubmit: jest.fn(), onCancel: jest.fn(), onOpenSettings: jest.fn(),
            getSkills: () => [],
            getAutoApply: () => stored,
            setAutoApply: (v) => { stored = v; },
        });
        const toggle = document.getElementById('autoApplyToggle');
        expect(toggle.checked).toBe(false);
        toggle.checked = true;
        toggle.dispatchEvent(new Event('change'));
        expect(stored).toBe(false);
        document.querySelector('[data-confirm="enable"]').click();
        await Promise.resolve();
        expect(window.confirm).not.toHaveBeenCalled();
        expect(stored).toBe(true);
    });

    test('cancelling auto-apply confirmation leaves it disabled and unpersisted', () => {
        window.confirm = jest.fn(() => false);
        let stored = false;
        initInputBar({
            onSubmit: jest.fn(), onCancel: jest.fn(), onOpenSettings: jest.fn(),
            getSkills: () => [],
            getAutoApply: () => stored,
            setAutoApply: (v) => { stored = v; },
        });
        const toggle = document.getElementById('autoApplyToggle');
        toggle.checked = true;
        toggle.dispatchEvent(new Event('change'));
        document.querySelector('[data-confirm="cancel"]').click();
        expect(window.confirm).not.toHaveBeenCalled();
        expect(toggle.checked).toBe(false);
        expect(stored).toBe(false);
    });
});

describe('IME composition and Enter handling', () => {
    test('does not submit while composing and submits after composition ends', () => {
        document.body.innerHTML = `
            <textarea id="chatInput"></textarea>
            <button id="sendBtn"></button>
            <div id="skillPicker" hidden></div>
            <div id="skillsMenu" hidden></div>
            <button id="addSkillBtn"></button>
            <input type="checkbox" id="autoApplyToggle">
            <button id="modelPill"></button>`;
        const onSubmit = jest.fn();
        initInputBar({ onSubmit, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
        const textarea = document.getElementById('chatInput');
        textarea.value = '中';
        textarea.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
        const composingEnter = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
        textarea.dispatchEvent(composingEnter);
        expect(composingEnter.defaultPrevented).toBe(false);
        expect(onSubmit).not.toHaveBeenCalled();

        textarea.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true }));
        const enter = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
        textarea.dispatchEvent(enter);
        expect(enter.defaultPrevented).toBe(true);
        expect(onSubmit).toHaveBeenCalledWith('中', []);
    });
});

describe('↑/↓ prompt history recall', () => {
    function setupWithHistory() {
        document.body.innerHTML = `
            <textarea id="chatInput"></textarea>
            <button id="sendBtn"></button>
            <div id="skillPicker" hidden></div>
            <div id="skillsMenu" hidden></div>
            <button id="addSkillBtn"></button>
            <input type="checkbox" id="autoApplyToggle">
            <button id="modelPill"></button>`;
        const onSubmit = jest.fn();
        initInputBar({ onSubmit, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn() });
        const textarea = document.getElementById('chatInput');
        const submit = (text) => {
            textarea.value = text;
            textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        };
        return { textarea, submit, onSubmit };
    }

    test('↑ recalls the newest prompt, ↑↑ walks back, ↓ returns forward', () => {
        const { textarea, submit } = setupWithHistory();
        submit('first prompt');
        submit('second prompt');

        textarea.value = '';
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
        expect(textarea.value).toBe('second prompt');
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
        expect(textarea.value).toBe('first prompt');
        // At the oldest entry ↑ is a no-op.
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
        expect(textarea.value).toBe('first prompt');

        textarea.setSelectionRange(textarea.value.length, textarea.value.length);
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
        expect(textarea.value).toBe('second prompt');
        // Past the newest entry the draft is restored.
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
        expect(textarea.value).toBe('');
    });

    test('an in-progress draft is preserved when recall passes it by', () => {
        const { textarea, submit } = setupWithHistory();
        submit('earlier prompt');

        textarea.value = 'my half-written draft';
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
        expect(textarea.value).toBe('earlier prompt');
        textarea.setSelectionRange(textarea.value.length, textarea.value.length);
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
        expect(textarea.value).toBe('my half-written draft');
    });

    test('consecutive duplicates collapse; a manual edit restarts from the newest', () => {
        const { textarea, submit } = setupWithHistory();
        submit('same text');
        submit('same text');

        textarea.value = '';
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
        expect(textarea.value).toBe('same text');

        // Manual edit exits navigation; ↑ starts from the newest again.
        textarea.value = 'same text edited';
        textarea.dispatchEvent(new Event('input', { bubbles: true }));
        textarea.setSelectionRange(0, 0);
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
        expect(textarea.value).toBe('same text');
    });

    test('↑ does not hijack caret movement inside multi-line text', () => {
        const { textarea, submit } = setupWithHistory();
        submit('earlier prompt');
        textarea.value = 'line1\nline2';
        textarea.setSelectionRange(textarea.value.length, textarea.value.length);
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', bubbles: true }));
        // Caret not at position 0 → the browser keeps default caret behavior
        // and the value is untouched.
        expect(textarea.value).toBe('line1\nline2');
    });
});

describe('file attachment chips', () => {
    function setupWithAttachments() {
        document.body.innerHTML = `
            <textarea id="chatInput"></textarea>
            <button id="sendBtn"></button>
            <div id="skillPicker" hidden></div>
            <div id="skillsMenu" hidden></div>
            <button id="addSkillBtn"></button>
            <input type="checkbox" id="autoApplyToggle">
            <button id="modelPill"></button>
            <div id="selectionPreview" hidden>
                <span id="selectionPreviewText"></span>
                <span id="selectionPreviewImages" hidden></span>
            </div>
            <div id="attachmentChips" hidden></div>
            <div id="inputError" role="alert" aria-live="assertive" hidden></div>
            <button id="attachBtn"></button>
            <input type="file" id="attachmentInput" hidden multiple>`;
        const onSubmit = jest.fn();
        const onLog = jest.fn();
        const bar = initInputBar({
            onSubmit, onCancel: jest.fn(), getSkills: () => [], onOpenSettings: jest.fn(), onLog,
        });
        const input = document.getElementById('attachmentInput');
        const pick = (files) => {
            Object.defineProperty(input, 'files', { value: files, configurable: true });
            input.dispatchEvent(new Event('change', { bubbles: true }));
        };
        return { bar, onSubmit, onLog, pick };
    }
    // FileReader-backed reads (jsdom Blobs lack .text()) settle a macrotask
    // late — give them a real timeout, not a 0ms tick.
    const flush = () => new Promise((r) => setTimeout(r, 30));

    test('picked text files render removable chips; submit passes and clears them', async () => {
        const { onSubmit, pick } = setupWithAttachments();
        pick([new File(['hello world'], 'notes.txt', { type: 'text/plain' })]);
        await flush();

        const chips = document.querySelectorAll('.attachment-chip');
        expect(chips).toHaveLength(1);
        expect(chips[0].querySelector('.attachment-chip-name').textContent).toBe('notes.txt');
        expect(document.getElementById('attachmentChips').hasAttribute('hidden')).toBe(false);

        const textarea = document.getElementById('chatInput');
        textarea.value = 'summarize this';
        textarea.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

        expect(onSubmit).toHaveBeenCalledTimes(1);
        const [text, attachments] = onSubmit.mock.calls[0];
        expect(text).toBe('summarize this');
        expect(attachments).toHaveLength(1);
        expect(attachments[0]).toMatchObject({ name: 'notes.txt', kind: 'text', text: 'hello world' });
        // Chips cleared after submit.
        expect(document.getElementById('attachmentChips').hasAttribute('hidden')).toBe(true);
    });

    test('attachments allow submitting with an empty input', async () => {
        const { onSubmit, pick } = setupWithAttachments();
        pick([new File(['body'], 'a.md')]);
        await flush();

        document.getElementById('chatInput').value = '   ';
        document.getElementById('sendBtn').click();
        expect(onSubmit).toHaveBeenCalledTimes(1);
        expect(onSubmit.mock.calls[0][1]).toHaveLength(1);
    });

    test('remove button drops a chip before sending', async () => {
        const { onSubmit, pick } = setupWithAttachments();
        pick([new File(['x'], 'a.txt'), new File(['y'], 'b.txt')]);
        await flush();
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(2);

        document.querySelectorAll('.attachment-chip-remove')[0].click();
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(1);
        expect(document.querySelector('.attachment-chip-name').textContent).toBe('b.txt');

        document.getElementById('chatInput').value = 'go';
        document.getElementById('sendBtn').click();
        expect(onSubmit.mock.calls[0][1]).toHaveLength(1);
        expect(onSubmit.mock.calls[0][1][0].name).toBe('b.txt');
    });

    test('unsupported and oversized files are rejected with a log entry', async () => {
        const { onLog, pick } = setupWithAttachments();
        pick([new File(['x'], 'archive.zip', { type: 'application/zip' })]);
        await flush();
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(0);
        expect(onLog).toHaveBeenCalledWith(expect.stringContaining('unsupported file type'), 'warning');
        const error = document.getElementById('inputError');
        expect(error.hidden).toBe(false);
        expect(error.textContent).toBe('archive.zip: unsupported file type (use text, image, .docx or .pdf).');

        const big = { name: 'huge.txt', type: 'text/plain', size: 10 * 1024 * 1024 + 1, text: async () => 'x' };
        pick([big]);
        await flush();
        expect(onLog).toHaveBeenCalledWith('huge.txt: 10.0 MB exceeds the 10.0 MB per-file limit.', 'warning');
        expect(error.hidden).toBe(false);
        expect(error.textContent).toBe('archive.zip: unsupported file type (use text, image, .docx or .pdf).\n'
            + 'huge.txt: 10.0 MB exceeds the 10.0 MB per-file limit.');

        document.getElementById('chatInput').dispatchEvent(new Event('input', { bubbles: true }));
        expect(error.hidden).toBe(false);
        expect(error.textContent).toContain('archive.zip');
    });

    test('parse failures show the filename only once in the inline error and log', async () => {
        const { pick, onLog } = setupWithAttachments();
        pick([{
            name: 'broken.txt', type: 'text/plain', size: 1,
            text: async () => { throw new Error('read failed'); },
        }]);
        // The chip appears as pending before the failure lands.
        expect(document.querySelectorAll('.attachment-chip-pending')).toHaveLength(1);
        await flush();
        const error = document.getElementById('inputError');
        expect(error.hidden).toBe(false);
        expect(error.textContent).toBe('broken.txt: read failed');
        expect(onLog).toHaveBeenCalledWith('broken.txt: read failed', 'error');
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(0);
    });

    test('picked files show a pending chip immediately, before parsing finishes', async () => {
        const { pick } = setupWithAttachments();
        let resolve;
        pick([
            { name: 'slow.txt', type: 'text/plain', size: 4, text: () => new Promise((r) => { resolve = r; }) },
            new File(['x'], 'fast.txt'),
        ]);

        // Chips render synchronously, before any parse settles.
        const chips = document.querySelectorAll('.attachment-chip');
        expect(chips).toHaveLength(2);
        for (const chip of chips) {
            expect(chip.classList.contains('attachment-chip-pending')).toBe(true);
            expect(chip.querySelector('.attachment-chip-spinner')).not.toBeNull();
            expect(chip.querySelector('.attachment-chip-status').textContent).toBe('Parsing…');
        }
        expect(chips[0].querySelector('.attachment-chip-name').textContent).toBe('slow.txt');
        expect(chips[1].querySelector('.attachment-chip-name').textContent).toBe('fast.txt');

        resolve('slow body');
        await flush();
        const settled = document.querySelectorAll('.attachment-chip');
        expect(settled).toHaveLength(2);
        for (const chip of settled) {
            expect(chip.classList.contains('attachment-chip-pending')).toBe(false);
            expect(chip.querySelector('.attachment-chip-spinner')).toBeNull();
            expect(chip.querySelector('.attachment-chip-status')).toBeNull();
        }
    });

    test('removing a pending chip discards its parse result', async () => {
        const { onSubmit, pick } = setupWithAttachments();
        let resolve;
        pick([{ name: 'slow.txt', type: 'text/plain', size: 4, text: () => new Promise((r) => { resolve = r; }) }]);
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(1);

        document.querySelector('.attachment-chip-remove').click();
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(0);

        resolve('late body');
        await flush();
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(0);

        document.getElementById('chatInput').value = 'go';
        document.getElementById('sendBtn').click();
        expect(onSubmit).toHaveBeenCalledTimes(1);
        expect(onSubmit.mock.calls[0][1]).toHaveLength(0);
    });

    test('image files render a thumbnail chip', async () => {
        const { pick } = setupWithAttachments();
        pick([new File([new Uint8Array([1, 2, 3])], 'pic.png', { type: 'image/png' })]);
        await flush();
        const thumb = document.querySelector('.attachment-chip-thumb');
        expect(thumb).not.toBeNull();
        expect(thumb.src).toMatch(/^data:image\/png;base64,/);
    });

    test('batch failures survive later success and ordinary typing until reset', async () => {
        const { bar, pick } = setupWithAttachments();
        pick([new File(['x'], 'bad.zip'), new File(['ok'], 'good.txt')]);
        await flush();
        const error = document.getElementById('inputError');
        expect(error.textContent).toContain('bad.zip');
        expect(document.querySelectorAll('.attachment-chip')).toHaveLength(1);
        document.getElementById('chatInput').dispatchEvent(new Event('input'));
        expect(error.hidden).toBe(false);
        bar.clearAttachments();
        expect(error.hidden).toBe(true);
    });

    test('pending parse blocks sending and cannot leak after a new session', async () => {
        const { bar, pick, onSubmit } = setupWithAttachments();
        let resolve;
        pick([{ name: 'old.txt', size: 4, type: 'text/plain', text: () => new Promise((r) => { resolve = r; }) }]);
        document.getElementById('chatInput').value = 'send';
        document.getElementById('sendBtn').click();
        expect(onSubmit).not.toHaveBeenCalled();
        bar.clearAttachments();
        pick([new File(['new'], 'new.txt')]);
        resolve('old');
        await flush();
        expect([...document.querySelectorAll('.attachment-chip-name')].map((el) => el.textContent)).toEqual(['new.txt']);
        document.getElementById('sendBtn').click();
        expect(onSubmit.mock.calls[0][1][0].name).toBe('new.txt');
    });

    test('clearAttachments empties the list', async () => {
        const { bar, pick } = setupWithAttachments();
        pick([new File(['x'], 'a.txt')]);
        await flush();
        bar.clearAttachments();
        expect(document.getElementById('attachmentChips').hasAttribute('hidden')).toBe(true);
    });
});

