/**
 * Panel-action enums — frozen string-tables for the panel-scoped button refactor.
 *
 * Three closed sets: prompt CATEGORY, button SCOPE, and the seven valid
 * ACTION strings emitted as `data-action` HTML attributes. Imported by the
 * dispatcher (taskpane.js), refactored handlers (Plan 04a), and Plan 02's
 * static-grep gate that validates HTML data-action strings against ACTION.*.
 * STYLE.md: "Enums for Fixed Values" + "No Magic Numbers or Strings".
 */

export const CATEGORY = Object.freeze({
  AMENDMENT: 'amendment',
  COMMENT:   'comment',
  SUMMARY:   'summary',
});

export const SCOPE = Object.freeze({
  SELECTION: 'selection',
  DOCUMENT:  'document',
});

export const ACTION = Object.freeze({
  AMEND_SELECTION:         'amend-selection',
  AMEND_COMMENT_SELECTION: 'amend-comment-selection',
  AMEND_DOCUMENT:          'amend-document',
  AMEND_COMMENT_DOCUMENT:  'amend-comment-document',
  COMMENT_SELECTION:       'comment-selection',
  COMMENT_DOCUMENT:        'comment-document',
  SUMMARY_DOCUMENT:        'summary-document',
});
