# Composing Word tasks

The planner describes user requirements separately from executable tasks. A requirement can need several native capabilities: document-wide space cleanup may involve body-text editing and table-cell editing, while table borders require the native table tools. `covers` therefore supports many tasks per requirement. Each reference must be valid and unique within its task; every requirement still needs coverage or an explicit unsupported explanation, and every action must have the correct document/answer effect. Invalid graphs, missing coverage, unavailable capabilities and contradictory unsupported declarations remain rejected.

An independent semantic review checks the original request, scope, contributions and dependencies against the same capability catalog used by planning. The catalog explicitly describes body-space cleanup, table-cell text, native borders, multi-table sessions and three-line tables. Repeated requirement IDs across tasks are not alone evidence of duplication. Conflicting redundant writes, invented actions and omissions are review failures. Space cleanup does not authorize deleting blank paragraphs or rewriting prose.

The reviewer receives its own trusted system message. Prior conversation is labeled request-context data, rather than replayed as assistant turns. A rejected review retains its findings, including missing/invented items and summary, for the bounded repair attempt. Contract and review errors appear in the activity log; final planning failure includes the concrete cause. An unapproved plan never falls back to executing an incomplete subset.

## Scope and execution

Tasks can specify `scope: "document" | "selection"`. An incidental selection does not narrow an explicit document request. Older task plans infer scope from their own instruction, then the original request, then current selection facts. Mixed-scope requests should specify scope on every task. Selection-only tasks without a selected target fail before execution.

Document-write dependencies remain staged until the preceding proposal is applied. The executor then reads current native document state for the next task; it does not consume an unapplied draft. Only tasks targeting the selection use the continuation selection guard. Document tasks can resume after a preceding edit changes an incidental selection.

For “全文优化格式，例如不正确的字体加粗，多余的空格，表格修改为三线格”, a supported decomposition is:

1. `format`, document scope: inspect and correct existing character/paragraph formatting.
2. `edit`, document scope, after formatting Apply: clean body-text spaces while preserving meaningful spacing, content and structure.
3. `table_management`, document scope, after body-edit Apply: re-read all native tables, clean cell spaces and set three-line borders. Both text tasks cover the space-cleanup requirement.

Three-line border instructions clear all existing table borders, restore the top and bottom rules, then add the header-bottom rule. Clearing only inside borders leaves existing left/right outer borders. Row overrides also need native row-border cleanup when present; structural table data and merged-cell restrictions still apply.

## Verification limits

Regression coverage includes the composite Chinese request, shared coverage, reviewer feedback repair, correct scopes with an active selection, chained proposal application, two native table drafts, removal of vertical borders, malformed plans and cancellation. These tests exercise contracts and mocked Word actions; they do not establish the exact reason for a previous model rejection or replace Mac Word runtime verification. The old generic rejection log omitted the review's actual findings, so its precise semantic objection cannot be recovered retrospectively.
