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

        items[0].click();
        expect(menu.hasAttribute('hidden')).toBe(true);
        expect(document.getElementById('chatInput').value).toBe('/mcp ');
    });

    test('auto-apply toggle initializes from config and persists changes', () => {
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
        expect(stored).toBe(true);
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
