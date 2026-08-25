/** @jest-environment jsdom */

/**
 * Proposal card tests: the staged-proposal UI. With no items the card keeps
 * its original all-or-nothing behavior; with items it renders an expandable
 * per-change list (checkbox + inline diff + locate link) and Apply forwards
 * only the checked ids.
 */

const { createProposalCard } = require('../src/taskpane/ui/proposal-card.js');

function makeItems() {
  return [
    { id: 'a', label: 'Section one', before: 'old text one', after: 'new text one', searchText: 'old text one' },
    { id: 'b', label: 'Section two', before: 'old text two', after: 'new text two', searchText: 'old text two' },
    { id: 'c', label: 'Insert at start', searchText: 'beginning' },
  ];
}

function makeCard(overrides = {}) {
  return createProposalCard({
    title: 'Proposed edits',
    countsText: '3 change op(s)',
    comment: null,
    onApply: jest.fn(async () => {}),
    onReject: jest.fn(),
    ...overrides,
  });
}

describe('createProposalCard', () => {
  test('without items there is no change list (original behavior)', () => {
    const card = makeCard();
    expect(card.el.querySelector('.proposal-card-changes')).toBeNull();
  });

  test('items render one checked checkbox row per change', () => {
    const card = makeCard({ items: makeItems() });
    const rows = card.el.querySelectorAll('.proposal-card-change');
    expect(rows).toHaveLength(3);
    expect(card.el.querySelector('summary').textContent).toBe('3 change(s)');
    const boxes = card.el.querySelectorAll('input[type="checkbox"]');
    expect([...boxes].every((b) => b.checked)).toBe(true);
    expect(rows[0].textContent).toContain('Section one');
  });

  test('before/after items render an inline diff with del/ins marks', () => {
    const card = makeCard({ items: makeItems() });
    const diffs = card.el.querySelectorAll('.diff-view');
    expect(diffs).toHaveLength(2); // third item has no before/after
    expect(diffs[0].querySelector('del').textContent).toBe('old');
    expect(diffs[0].querySelector('ins').textContent).toBe('new');
  });

  test('locate button fires onLocate with the item searchText', () => {
    const onLocate = jest.fn();
    const card = makeCard({ items: makeItems(), onLocate });
    card.el.querySelectorAll('.proposal-card-locate')[1].click();
    expect(onLocate).toHaveBeenCalledWith('old text two');
  });

  test('apply forwards only the checked ids', async () => {
    const onApply = jest.fn(async () => {});
    const card = makeCard({ items: makeItems(), onApply });
    const boxes = card.el.querySelectorAll('input[type="checkbox"]');
    boxes[1].checked = false;
    card.el.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(onApply).toHaveBeenCalledWith(['a', 'c']);
  });

  test('apply without items forwards undefined', async () => {
    const onApply = jest.fn(async () => {});
    const card = makeCard({ onApply });
    card.el.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(onApply).toHaveBeenCalledWith(undefined);
  });

  test('unchecking everything disables Apply; re-checking re-enables it', () => {
    const card = makeCard({ items: makeItems() });
    const boxes = card.el.querySelectorAll('input[type="checkbox"]');
    const applyBtn = card.el.querySelector('.btn-primary');
    boxes.forEach((b) => { b.checked = false; b.dispatchEvent(new Event('change')); });
    expect(applyBtn.disabled).toBe(true);
    boxes[0].checked = true;
    boxes[0].dispatchEvent(new Event('change'));
    expect(applyBtn.disabled).toBe(false);
  });

  test('markApplied reports the partial-apply count', () => {
    const card = makeCard({ items: makeItems() });
    const boxes = card.el.querySelectorAll('input[type="checkbox"]');
    boxes[2].checked = false;
    card.el.querySelector('.btn-primary').click();
    card.markApplied();
    expect(card.el.querySelector('.proposal-card-status').textContent)
      .toBe('Applied 2 of 3 change(s) as tracked changes.');
  });

  test('markApplied keeps the original line without items', () => {
    const card = makeCard();
    card.el.querySelector('.btn-primary').click();
    card.markApplied();
    expect(card.el.querySelector('.proposal-card-status').textContent).toBe('Applied as tracked changes.');
  });

  test('markError re-enables Apply (unless everything is unchecked)', async () => {
    // Mirrors the real consumer pattern (conversation.js): onApply catches
    // the error itself and settles the card via markError.
    let card;
    const onApply = jest.fn(async () => {
      try {
        throw new Error('boom');
      } catch (e) {
        card.markError(e.message);
      }
    });
    card = makeCard({ items: makeItems(), onApply });
    card.el.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(card.el.querySelector('.btn-primary').disabled).toBe(false);
    expect(card.el.querySelector('.proposal-card-status').textContent).toBe('Apply failed: boom');
  });

  test('markWarning settles with the warning class and keeps buttons disabled', () => {
    const card = makeCard({ items: makeItems() });
    card.markWarning('Nothing applied: chunk skipped');
    expect(card.el.classList.contains('proposal-warning')).toBe(true);
    expect(card.el.classList.contains('proposal-applied')).toBe(false);
    const status = card.el.querySelector('.proposal-card-status');
    expect(status.textContent).toBe('Nothing applied: chunk skipped');
    expect(status.style.display).not.toBe('none');
    expect(card.el.querySelector('.btn-primary').disabled).toBe(true);
    expect(card.el.querySelector('.btn-secondary').disabled).toBe(true);
  });
});
