# Composing Word tasks

The planner describes user requirements separately from executable tasks. A requirement can need several native capabilities: document-wide space cleanup may involve body-text editing and table-cell editing, while table borders require the native table tools. `covers` therefore supports many tasks per requirement. Each reference must be valid and unique within its task; every requirement still needs coverage or an explicit unsupported explanation, and every action must have the correct document/answer effect. Invalid graphs, missing coverage, unavailable capabilities and contradictory unsupported declarations remain rejected.

An independent semantic review checks the original request, scope, contributions and dependencies against the same capability catalog used by planning. The catalog explicitly describes body-space cleanup, table-cell text, native borders, multi-table sessions and three-line tables. Repeated requirement IDs across tasks are not alone evidence of duplication. Conflicting redundant writes, invented actions and omissions are review failures. Space cleanup does not authorize deleting blank paragraphs or rewriting prose.

The reviewer receives its own trusted system message. Prior conversation is labeled request-context data, rather than replayed as assistant turns. A rejected review retains its findings, including missing/invented items and summary, for the bounded repair attempt. Contract and review errors appear in the activity log; final planning failure includes the concrete cause. An unapproved plan never falls back to executing an incomplete subset.

## Scope and execution

Tasks can specify `scope: "document" | "selection"`. An incidental selection does not narrow an explicit document request. Older task plans infer scope from their own instruction, then the original request, then current selection facts. Mixed-scope requests should specify scope on every task. Selection-only tasks without a selected target fail before execution.

Document-write dependencies remain staged until the preceding proposal is applied. The executor then reads current native document state for the next task; it does not consume an unapplied draft. Only tasks targeting the selection use the continuation selection guard. Document tasks can resume after a preceding edit changes an incidental selection.

The executor also enforces this application boundary for native writes when the model omits dependencies. It stages one write at a time while independent read-only answers can proceed. Independent failed tasks do not become artificial prerequisites. Failed results remain terminal during Apply continuation; blocked tasks are reconsidered without silently retrying failed work. Compatibility task IDs stay stable across application and resumption. Confirmed preservation of necessary structural empty paragraphs permits continuation; unreadable or partially applied changes still stop it. Task failures report their individual causes and the actual number of pending proposals. Stop prevents previously queued proposals from auto-applying when a later task is cancelled.

For “全文优化格式，例如不正确的字体加粗，多余的空格，表格修改为三线格”, a supported decomposition is:

1. `format`, document scope: inspect and correct existing character/paragraph formatting.
2. `edit`, document scope, after formatting Apply: clean body-text spaces while preserving meaningful spacing, content and structure.
3. `table_management`, document scope, after body-edit Apply: re-read all native tables, clean cell spaces and set three-line borders. Both text tasks cover the space-cleanup requirement.

Three-line border instructions clear all existing table borders, restore the top and bottom rules, then add the header-bottom rule. Clearing only inside borders leaves existing left/right outer borders. Row overrides also need native row-border cleanup when present; structural table data and merged-cell restrictions still apply.

## Verification limits

Whole-document formatting uses the native `body.getRange('Whole')` at both preparation and application, rather than recovering a bookmark spanning tables and body boundaries. Exact text and verified native OOXML fingerprints still reject stale proposals. Selection operations retain strict bookmark recovery and never widen to the whole document.

Space-only body tasks use a conservative scan of readable visible non-table text. A clean source produces a verified no-op before model calls or bookmark allocation. Candidate spaces go to the model; returned edits may delete ASCII spaces but must preserve words, punctuation, native paragraph structure, manual line breaks, tabs and non-breaking spaces. Boundary-space changes use exact comparisons and character-level writes. The scan excludes tables, empty-paragraph cleanup, hidden revision/field code and non-text objects; it does not prove whether equation or alignment spacing is meaningful. Unchanged candidate text remains unverified rather than being reported as completed.

Required lazy modules are preloaded after planning and before expensive document tasks. Script download failures receive two bounded retries; native actions are never replayed by the loader. A caller can cancel its wait without cancelling another caller's shared download. Final failure gives a connection/reopen-pane instruction without silently reloading and losing pending proposals. Production script download attempts time out after 30 seconds. All publishing paths preserve immutable hashed JavaScript chunks and licenses for 30 days so open panes can finish using an earlier runtime. The reported `agent-actions.875752d6.js` URL returned HTTP 200 during diagnosis; the supplied run does not establish whether its failure was transient delivery or a stale deployment.

Regression coverage includes the composite Chinese request, shared coverage, reviewer feedback repair, correct scopes with an active selection, chained proposal application, two native table drafts, removal of vertical borders, malformed plans and cancellation. These tests exercise contracts and mocked Word actions; they do not establish the exact reason for a previous model rejection or replace Mac Word runtime verification. The old generic rejection log omitted the review's actual findings, so its precise semantic objection cannot be recovered retrospectively.
