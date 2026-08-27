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

    test('empty content and legacy strings still work', () => {
        const bar = setupDom();
        bar.setSelectionPreview({ text: '', images: [], totalImages: 0 });
        expect(document.getElementById('selectionPreview').hasAttribute('hidden')).toBe(true);

        // Legacy string form (text-only) renders as before.
        bar.setSelectionPreview('plain string');
        expect(document.getElementById('selectionPreviewText').textContent).toBe('plain string');
    });
});
