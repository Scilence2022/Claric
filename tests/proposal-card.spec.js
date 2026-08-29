/** @jest-environment jsdom */

/**
 * Proposal card tests: the staged-proposal UI. With no items the card keeps
 * its original all-or-nothing behavior; with items it renders an expandable
 * per-change list (checkbox + inline diff + locate link) and Apply forwards
 * only the checked ids.
 */

const { createProposalCard, renderTablePreview, sanitizeTablePreview } = require('../src/taskpane/ui/proposal-card.js');

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
  test('previewSvg renders inline sanitized SVG (no img, no active content)', () => {
    // Regression: SVG previews used to ship as base64 data-URL <img>s, which
    // decode fine in Chromium but fail on some hosts (WKWebView taskpane).
    // The inline variant renders the sanitized markup directly.
    const card = makeCard({
      previewSvg: '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80">'
        + '<rect width="120" height="80" fill="#eef"/><circle cx="60" cy="40" r="20" fill="#2e9e6b"/>'
        + '<script>alert(1)</' + 'script></svg>',
    });
    const holder = card.el.querySelector('.proposal-card-preview-svg');
    expect(holder).not.toBeNull();
    expect(holder.querySelector('svg')).not.toBeNull();
    // Render-time sanitization strips active content even if staged markup
    // somehow carried it.
    expect(holder.querySelector('script')).toBeNull();
    expect(card.el.querySelector('img.proposal-card-preview')).toBeNull();
  });

  test('without items there is no change list (original behavior)', () => {    const card = makeCard();
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
    expect(onApply).toHaveBeenCalledWith(['a', 'c'], expect.objectContaining({
      signal: expect.anything(),
      onChunkApplied: expect.any(Function),
    }));
  });

  test('apply without items forwards undefined', async () => {
    const onApply = jest.fn(async () => {});
    const card = makeCard({ onApply });
    card.el.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    expect(onApply).toHaveBeenCalledWith(undefined, expect.anything());
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

  test('markItemApplied dims + disables the row and adds a status tag', () => {
    const card = makeCard({ items: makeItems() });
    const rows = card.el.querySelectorAll('.proposal-card-change');
    card.markItemApplied('a', { applied: true });

    expect(card.el.querySelectorAll('input[type="checkbox"]')[0].disabled).toBe(true);
    expect(card.el.querySelectorAll('input[type="checkbox"]')[0].checked).toBe(true);
    expect(rows[0].classList.contains('proposal-card-change-done')).toBe(true);
    expect(rows[0].querySelector('.proposal-card-change-status').textContent).toBe('applied');
    // Unrelated rows untouched.
    expect(rows[1].classList.contains('proposal-card-change-done')).toBe(false);
  });

  test('markItemApplied tags error / skipped outcomes distinctly', () => {
    const card = makeCard({ items: makeItems() });
    const rows = card.el.querySelectorAll('.proposal-card-change');
    card.markItemApplied('a', { error: true });
    card.markItemApplied('b', { skipped: true });
    expect(rows[0].querySelector('.proposal-card-change-status').textContent).toBe('error');
    expect(rows[0].classList.contains('proposal-card-change-error')).toBe(true);
    expect(rows[1].querySelector('.proposal-card-change-status').textContent).toBe('skipped');
  });

  test('setPaused re-enables Apply as "Continue applying"', () => {
    const card = makeCard({ items: makeItems() });
    card.markItemApplied('a', { applied: true });
    card.setPaused('Paused — 1 of 3 applied.');

    const applyBtn = card.el.querySelector('.btn-primary');
    expect(applyBtn.disabled).toBe(false);
    expect(applyBtn.textContent).toBe('Continue applying');
    expect(card.el.querySelector('.proposal-card-status').textContent).toContain('Paused');
    expect(card.el.classList.contains('proposal-paused')).toBe(true);
  });

  test('continue (resume) forwards only the still-pending ids', async () => {
    const onApply = jest.fn(async () => {});
    const card = makeCard({ items: makeItems(), onApply });
    card.markItemApplied('a', { applied: true });
    card.setPaused('Paused — 1 of 3 applied.');

    card.el.querySelector('.btn-primary').click();
    await new Promise((r) => setTimeout(r, 0));
    // Applied 'a' is excluded from the resumed selection.
    expect(onApply).toHaveBeenCalledWith(['b', 'c'], expect.anything());
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

  test('reject click settles through the public markRejected API', () => {
    const card = makeCard();
    const spy = jest.spyOn(card, 'markRejected');
    card.el.querySelector('.btn-secondary').click();
    expect(spy).toHaveBeenCalledTimes(1);
    expect(card.el.classList.contains('proposal-rejected')).toBe(true);
    expect(card.el.querySelector('.proposal-card-status').textContent)
      .toBe('Rejected — no changes were made.');
  });
});

describe('table preview', () => {
  const preview = {
    rows: [['Name', 'Value'], ['Alpha', '1']],
    headerRowCount: 1,
    style: 'tableGrid',
    position: 'end',
  };

  test('createProposalCard renders a read-only matrix with header semantics', () => {
    const card = makeCard({ tablePreview: preview });
    const table = card.el.querySelector('.proposal-card-table');
    expect(table).not.toBeNull();
    expect(table.querySelectorAll('thead th')).toHaveLength(2);
    expect(table.querySelectorAll('tbody td')).toHaveLength(2);
    expect(table.querySelector('thead th').getAttribute('scope')).toBe('col');
    expect(card.el.querySelector('.proposal-card-table-meta').textContent)
      .toContain('Dimensions: 2 × 2');
  });

  test('cell text is assigned via textContent (markup cannot inject)', () => {
    const card = makeCard({
      tablePreview: { rows: [['<img src=x onerror=alert(1)>']], headerRowCount: 0 },
    });
    const cell = card.el.querySelector('.proposal-card-table td');
    expect(cell.textContent).toBe('<img src=x onerror=alert(1)>');
    expect(card.el.querySelector('.proposal-card-table img')).toBeNull();
  });

  test('createProposalCard without tablePreview renders no preview block', () => {
    const card = makeCard();
    expect(card.el.querySelector('.proposal-card-table-preview')).toBeNull();
  });

  test('sanitizeTablePreview normalizes ragged/oversized/invalid input', () => {
    expect(sanitizeTablePreview(null)).toBeNull();
    expect(sanitizeTablePreview({ rows: 'nope' })).toBeNull();

    const normalized = sanitizeTablePreview({
      rows: [['a', 'b'], ['c'], null],
      headerRowCount: 5,
      style: 'x'.repeat(200),
      position: 42,
    });
    expect(normalized.rows).toEqual([['a', 'b'], ['c', ''], ['', '']]);
    expect(normalized.headerRowCount).toBe(3);
    expect(normalized.style).toHaveLength(80);
    expect(normalized).not.toHaveProperty('position');
    expect(normalized.truncated).toBe(true);
  });

  test('renderTablePreview marks empty input without throwing', () => {
    const el = renderTablePreview({ rows: [], headerRowCount: 0 });
    expect(el.querySelector('.proposal-card-table-empty').textContent)
      .toBe('No table cells in preview.');
  });
});

describe('cross-card apply mutex and busy wiring', () => {
  const flush = () => new Promise((r) => setTimeout(r, 0));

  test('a second card refuses to apply while another apply is in flight', async () => {
    let releaseA;
    const cardA = makeCard({ onApply: () => new Promise((resolve) => { releaseA = resolve; }) });
    const onApplyB = jest.fn(async () => {});
    const cardB = makeCard({ onApply: onApplyB });

    cardA.el.querySelector('.btn-primary').click();
    await flush(); // A's onApply now pending

    cardB.el.querySelector('.btn-primary').click();
    await flush();
    expect(onApplyB).not.toHaveBeenCalled();
    expect(cardB.el.querySelector('.proposal-card-status').textContent)
      .toMatch(/Another proposal/);

    releaseA();
    await flush(); await flush(); // A settles
    // After A settles, B can apply again.
    cardB.el.querySelector('.btn-primary').click();
    await flush();
    expect(onApplyB).toHaveBeenCalledTimes(1);
  });

  test('setApplyBusy flips true→false around the apply', async () => {
    const setApplyBusy = jest.fn();
    const card = makeCard({ setApplyBusy });
    card.el.querySelector('.btn-primary').click();
    await flush();
    expect(setApplyBusy.mock.calls[0][0]).toBe(true);
    await flush();
    expect(setApplyBusy.mock.calls[setApplyBusy.mock.calls.length - 1][0]).toBe(false);
  });
});

describe('applyAll (auto-apply path)', () => {
    const flush = () => new Promise((r) => setTimeout(r, 0));

    test('applies all checked changes without a button click', async () => {
        const onApply = jest.fn(async () => {});
        const card = makeCard({ items: makeItems(), onApply });
        card.applyAll();
        await flush();
        expect(onApply).toHaveBeenCalledWith(['a', 'b', 'c'], expect.anything());
        // Without a terminal settle (markApplied etc.) the button stays
        // disabled — the onApply handler owns the terminal state, exactly
        // as with a manual click.
    });

    test('respects the cross-card mutex', async () => {
        let release;
        const cardA = makeCard({ onApply: () => new Promise((resolve) => { release = resolve; }) });
        const onApplyB = jest.fn(async () => {});
        const cardB = makeCard({ onApply: onApplyB });

        cardA.applyAll();
        await flush();
        cardB.applyAll();
        await flush();
        expect(onApplyB).not.toHaveBeenCalled();
        release();
        await flush(); await flush();
    });

    test('isBlocked refusal aborts before onApply and before registering the controller', async () => {
        // The C-1 regression: a card apply during an in-flight run used to
        // hijack the run's controller (uncancellable run) and release the
        // busy flags early (concurrent pipelines). The pre-flight guard must
        // refuse before either happens.
        const registerController = jest.fn();
        const setApplyBusy = jest.fn();
        const onApply = jest.fn(async () => {});
        const card = makeCard({
            items: makeItems(),
            onApply,
            registerController,
            setApplyBusy,
            isBlocked: () => 'A run is currently processing — wait for it to finish before applying.',
        });

        card.applyAll();
        await new Promise((r) => setTimeout(r, 0));

        expect(onApply).not.toHaveBeenCalled();
        expect(registerController).not.toHaveBeenCalled();
        expect(setApplyBusy).not.toHaveBeenCalled();
        const status = card.el.querySelector('.proposal-card-status');
        expect(status.textContent).toContain('A run is currently processing');
    });

    test('isBlocked null allows the apply (guard is a no-op when idle)', async () => {
        const onApply = jest.fn(async () => {});
        const card = makeCard({ items: makeItems(), onApply, isBlocked: () => null });

        card.applyAll();
        await new Promise((r) => setTimeout(r, 0));
        expect(onApply).toHaveBeenCalledTimes(1);
    });
});
