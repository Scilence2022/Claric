const { CATEGORY, SCOPE, ACTION } = require('../src/lib/panel-actions.js');

describe('panel-actions enums', () => {
  test('CATEGORY is frozen and has the three expected entries', () => {
    expect(Object.isFrozen(CATEGORY)).toBe(true);
    expect(CATEGORY).toEqual({
      AMENDMENT: 'amendment',
      COMMENT:   'comment',
      SUMMARY:   'summary',
    });
  });

  test('SCOPE is frozen and has the two expected entries', () => {
    expect(Object.isFrozen(SCOPE)).toBe(true);
    expect(SCOPE).toEqual({
      SELECTION: 'selection',
      DOCUMENT:  'document',
    });
  });

  test('ACTION is frozen and has exactly 7 entries (one per RESEARCH.md Pattern 3 row)', () => {
    expect(Object.isFrozen(ACTION)).toBe(true);
    expect(Object.values(ACTION).length).toBe(7);
    expect(ACTION).toEqual({
      AMEND_SELECTION:         'amend-selection',
      AMEND_COMMENT_SELECTION: 'amend-comment-selection',
      AMEND_DOCUMENT:          'amend-document',
      AMEND_COMMENT_DOCUMENT:  'amend-comment-document',
      COMMENT_SELECTION:       'comment-selection',
      COMMENT_DOCUMENT:        'comment-document',
      SUMMARY_DOCUMENT:        'summary-document',
    });
  });

  test('all ACTION strings are kebab-case (matches HTML data-action convention)', () => {
    Object.values(ACTION).forEach((s) => {
      expect(s).toMatch(/^[a-z]+(-[a-z]+)+$/);
    });
  });
});
