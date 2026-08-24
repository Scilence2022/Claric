/** @jest-environment jsdom */

/**
 * Panel-button enable-rule tests (AC-05 / AC-06 / AC-07).
 *
 * Targets functions that Plan 03 will export from
 * `src/taskpane/taskpane.js`:
 *
 *   - getPanelButtonState(category) -> { enabled, withCommentEnabled }
 *   - updatePanelButtons(category)  -> mutates DOM `disabled` flags
 *
 * Until Plan 03 lands, those exports are missing. These tests therefore
 * FAIL LOUDLY with messages naming `getPanelButtonState` /
 * `updatePanelButtons`. Failing-by-design: Plan 03's <verify> block
 * will reference these test names and turn them green.
 *
 * The Nyquist contract in 05.1-VALIDATION.md requires every later
 * plan's <verify> to point at real, currently-failing tests. Do NOT
 * convert these to test.todo — todos pass quietly and would lie to
 * the sampling rate.
 */

const { CATEGORY, ACTION } = require('../src/lib/panel-actions.js');

// Optimistic import: Plan 03 will export these. Until then, capture
// whatever the import fails with so each test can produce a clear
// "missing export" assertion failure (rather than a silent crash).
let getPanelButtonState;
let updatePanelButtons;
let importErr = null;
try {
  ({ getPanelButtonState, updatePanelButtons } = require('../src/taskpane/taskpane.js'));
} catch (e) {
  importErr = e;
}

// Mockable promptManager: Plan 03's getPanelButtonState reads the
// active prompt + commentInstructions textarea to compute enable state.
const mockPromptManager = {
  getActivePrompt: jest.fn(),
};

const FIXTURE_HTML = `
  <div role="tabpanel" id="panel-amendment">
    <button data-panel="${CATEGORY.AMENDMENT}" data-action="${ACTION.AMEND_SELECTION}">Amend Selection &rarr;</button>
    <button data-panel="${CATEGORY.AMENDMENT}" data-action="${ACTION.AMEND_COMMENT_SELECTION}">Amend &amp; Comment Selection &rarr;</button>
    <button data-panel="${CATEGORY.AMENDMENT}" data-action="${ACTION.AMEND_DOCUMENT}">Amend Document &rarr;</button>
    <button data-panel="${CATEGORY.AMENDMENT}" data-action="${ACTION.AMEND_COMMENT_DOCUMENT}">Amend &amp; Comment Document &rarr;</button>
    <textarea id="commentInstructions"></textarea>
  </div>
`;

beforeEach(() => {
  document.body.innerHTML = FIXTURE_HTML;
  mockPromptManager.getActivePrompt.mockReset();
});

describe('panel buttons (AC-05 / AC-06 / AC-07)', () => {
  test('AC-05: no active Amendment prompt -> all 4 Amendment buttons disabled', () => {
    expect(importErr).toBeNull();
    expect(typeof getPanelButtonState).toBe('function');
    expect(typeof updatePanelButtons).toBe('function');

    mockPromptManager.getActivePrompt.mockReturnValue(null);

    const state = getPanelButtonState(CATEGORY.AMENDMENT, { promptManager: mockPromptManager });
    expect(state).toEqual({ enabled: false, withCommentEnabled: false });

    updatePanelButtons(CATEGORY.AMENDMENT, { promptManager: mockPromptManager });

    const buttons = document.querySelectorAll(
      `button[data-panel="${CATEGORY.AMENDMENT}"][data-action]`
    );
    expect(buttons.length).toBe(4);
    buttons.forEach((btn) => {
      expect(btn.disabled).toBe(true);
    });
  });

  test('AC-06: Amendment prompt active + empty commentInstructions -> non-comment buttons enable, "& Comment" buttons stay disabled', () => {
    expect(importErr).toBeNull();
    expect(typeof getPanelButtonState).toBe('function');
    expect(typeof updatePanelButtons).toBe('function');

    mockPromptManager.getActivePrompt.mockImplementation((cat) =>
      cat === CATEGORY.AMENDMENT ? { id: 'amend-1', name: 'Default amend' } : null
    );

    const ta = document.getElementById('commentInstructions');
    ta.value = ''; // explicit: empty

    const state = getPanelButtonState(CATEGORY.AMENDMENT, { promptManager: mockPromptManager });
    expect(state).toEqual({ enabled: true, withCommentEnabled: false });

    updatePanelButtons(CATEGORY.AMENDMENT, { promptManager: mockPromptManager });

    const plain = document.querySelector(
      `button[data-panel="${CATEGORY.AMENDMENT}"][data-action="${ACTION.AMEND_SELECTION}"]`
    );
    const plainDoc = document.querySelector(
      `button[data-panel="${CATEGORY.AMENDMENT}"][data-action="${ACTION.AMEND_DOCUMENT}"]`
    );
    const withCommentSel = document.querySelector(
      `button[data-panel="${CATEGORY.AMENDMENT}"][data-action="${ACTION.AMEND_COMMENT_SELECTION}"]`
    );
    const withCommentDoc = document.querySelector(
      `button[data-panel="${CATEGORY.AMENDMENT}"][data-action="${ACTION.AMEND_COMMENT_DOCUMENT}"]`
    );

    expect(plain.disabled).toBe(false);
    expect(plainDoc.disabled).toBe(false);
    expect(withCommentSel.disabled).toBe(true);
    expect(withCommentDoc.disabled).toBe(true);
  });

  test('AC-07: Amendment prompt active + filled commentInstructions -> all 4 buttons enable', () => {
    expect(importErr).toBeNull();
    expect(typeof getPanelButtonState).toBe('function');
    expect(typeof updatePanelButtons).toBe('function');

    mockPromptManager.getActivePrompt.mockImplementation((cat) =>
      cat === CATEGORY.AMENDMENT ? { id: 'amend-1', name: 'Default amend' } : null
    );

    const ta = document.getElementById('commentInstructions');
    ta.value = 'please be polite';

    const state = getPanelButtonState(CATEGORY.AMENDMENT, { promptManager: mockPromptManager });
    expect(state).toEqual({ enabled: true, withCommentEnabled: true });

    updatePanelButtons(CATEGORY.AMENDMENT, { promptManager: mockPromptManager });

    const buttons = document.querySelectorAll(
      `button[data-panel="${CATEGORY.AMENDMENT}"][data-action]`
    );
    expect(buttons.length).toBe(4);
    buttons.forEach((btn) => {
      expect(btn.disabled).toBe(false);
    });
  });
});
