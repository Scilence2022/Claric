/**
 * runAction dispatcher tests (AC-08, AC-09 + dispatch-table-completeness).
 *
 * Targets exports that Plan 03 will add to `src/taskpane/taskpane.js`:
 *
 *   - createDispatcher(deps)  -> dispatch({category, scope, withComment})
 *   - ROUTES (named export)   -> frozen object keyed by
 *                                "category:scope:plain|withComment"
 *
 * Three INTENTIONALLY FAILING tests — Plan 03 turns them green.
 * Failing-by-design per the Nyquist contract (05.1-VALIDATION.md).
 *
 * NO jsdom pragma: the dispatcher is pure JS; no DOM required.
 *
 * STYLE.md: "Dispatch Over If/Else Chains" — Test 3 is the
 * dispatch-table-completeness gate. If Plan 03 ships an if/else
 * implementation instead of a frozen ROUTES table, Test 3 fails.
 */

const { CATEGORY, SCOPE } = require('../src/lib/panel-actions.js');

let createDispatcher;
let ROUTES;
let importErr = null;
try {
  // eslint-disable-next-line global-require
  ({ createDispatcher, ROUTES } = require('../src/taskpane/taskpane.js'));
} catch (e) {
  importErr = e;
}

const EXPECTED_ROUTE_KEYS = [
  'amendment:selection:plain',
  'amendment:selection:withComment',
  'amendment:document:plain',
  'amendment:document:withComment',
  'comment:selection:plain',
  'comment:document:plain',
  'summary:document:plain',
];

describe('runAction dispatcher', () => {
  test('AC-08: amendment selection routes by arg even when getActiveMode would say "summary" (and getActiveMode is NEVER called)', () => {
    expect(importErr).toBeNull();
    expect(typeof createDispatcher).toBe('function');

    const amendmentSelectionSpy = jest.fn();
    const processDocSpy         = jest.fn();
    const summarySpy            = jest.fn();
    const commentSpy            = jest.fn();
    const mockPromptManager = {
      // If the dispatcher reads getActiveMode (the bug AC-08 forbids),
      // it would see 'summary' and route to the wrong handler.
      getActiveMode:    jest.fn(() => 'summary'),
      getActivePrompt:  jest.fn(() => ({ id: 'amend-1' })),
    };

    const dispatch = createDispatcher({
      promptManager:           mockPromptManager,
      handleReviewSelection:   amendmentSelectionSpy,
      handleProcessDocument:   processDocSpy,
      handleSummaryGeneration: summarySpy,
      fireCommentRequest:      commentSpy,
      isProcessingDocRef:      () => false,
      addLog:                  jest.fn(),
    });

    dispatch({
      category:    CATEGORY.AMENDMENT,
      scope:       SCOPE.SELECTION,
      withComment: false,
    });

    expect(amendmentSelectionSpy).toHaveBeenCalledTimes(1);
    expect(mockPromptManager.getActiveMode).not.toHaveBeenCalled();
    expect(summarySpy).not.toHaveBeenCalled();
    expect(processDocSpy).not.toHaveBeenCalled();
    expect(commentSpy).not.toHaveBeenCalled();
  });

  test('AC-09: summary:document routes to handleSummaryGeneration', () => {
    expect(importErr).toBeNull();
    expect(typeof createDispatcher).toBe('function');

    const amendmentSelectionSpy = jest.fn();
    const processDocSpy         = jest.fn();
    const summarySpy            = jest.fn();
    const commentSpy            = jest.fn();
    const mockPromptManager = {
      getActiveMode:    jest.fn(() => 'amendment'),
      getActivePrompt:  jest.fn(() => ({ id: 'sum-1' })),
    };

    const dispatch = createDispatcher({
      promptManager:           mockPromptManager,
      handleReviewSelection:   amendmentSelectionSpy,
      handleProcessDocument:   processDocSpy,
      handleSummaryGeneration: summarySpy,
      fireCommentRequest:      commentSpy,
      isProcessingDocRef:      () => false,
      addLog:                  jest.fn(),
    });

    dispatch({
      category:    CATEGORY.SUMMARY,
      scope:       SCOPE.DOCUMENT,
      withComment: false,
    });

    expect(summarySpy).toHaveBeenCalledTimes(1);
    expect(amendmentSelectionSpy).not.toHaveBeenCalled();
    expect(processDocSpy).not.toHaveBeenCalled();
    expect(commentSpy).not.toHaveBeenCalled();
  });

  test('STYLE.md "Dispatch Over If/Else": ROUTES is a frozen object with exactly 7 keys, each mapping to a function', () => {
    expect(importErr).toBeNull();
    expect(ROUTES).toBeDefined();
    expect(typeof ROUTES).toBe('object');

    const expected = new Set(EXPECTED_ROUTE_KEYS);
    const actual   = new Set(Object.keys(ROUTES));
    expect(actual).toEqual(expected);

    // No keys beyond the 7 valid tuples (catches accidental future
    // additions without spec updates).
    expect(Object.keys(ROUTES).length).toBe(7);

    // Every value is a callable handler.
    expect(Object.values(ROUTES).every((v) => typeof v === 'function')).toBe(true);
  });
});
