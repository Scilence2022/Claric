# Word diff performance analysis

## Scope and evidence

Analysis and fixes are based on main commit `6f1c41c` (including PRs #202 and
#203). No manuscript or live Word document was modified during this work.

The reported sequence is unusually useful for diagnosis:

1. Full-document polishing leaves scrolling visibly slow.
2. Clearing Undo, closing the Claric pane, and selecting No Markup do not
   resolve it. Save As alone also does not resolve it.
3. Reopening the saved document restores responsiveness without restarting
   Word. The same Word process remains running.
4. OS samples before reopening repeatedly show 80–100% CPU in Word, largely
   on its main thread. After reopening, sampled CPU falls to approximately
   10–23%, with most main-thread samples waiting for events. Process memory
   remains roughly 6.2–6.5 GiB. These are observations, not an identical
   automated scrolling benchmark.

This supports transient native document state as a contributor. It does not
identify a particular Word cache, revision index, piece table, or memory leak:
the dominant native Word frames lack public symbols. A low-CPU add-in could
still have caused expensive native work through earlier writes.

## What is actually third-party code?

`src/lib/word-diff/NOTICE` pins the vendored source to
[office-word-diff, commit 81315f2](https://github.com/yuch85/office-word-diff/tree/81315f2fae23ba8cf9f40c13bdfd17d0e444d137).
This analysis compared that commit's sources with Claric, rather than assuming
that the directory still runs the unmodified upstream library.

The upstream token strategy searches every token, takes the first match,
deletes tokens individually, and resets the whole range to the original text
before falling back after an error. Those patterns can multiply native work
and can be unsafe around existing revisions. Claric already replaces these
patterns with selective occurrence-aware searches, contiguous deletion runs,
revision-aware validation, and refusal to retry after a possible write.
They are not newly discovered active bugs in the current token path.

The JavaScript diff engine produces an edit plan and finishes; it has no
persistent timer that would keep scrolling slow after the pane closes. Native
Word work induced by the plan is a more plausible link to this symptom.

## The effective full-document path

```text
applyChunkResults
  -> paragraph preflight (nativeTargets: true)
  -> configured token / sentence / character strategy
       -> tryRevisionDiff
            -> current text + OOXML and immutable-object validation
            -> edit planning and verified native range location
            -> tracked text mutations + read-back
  -> paragraph and section verification
```

`nativeTargets: true` means that an OOXML-capable host uses Claric's native
revision-aware path even for a paragraph with no earlier revisions. The
vendored token implementation below the early return is therefore not the
normal full-document write engine on such hosts.

## Confirmed problems and fixes

### Strategy granularity was lost at the early return

All three entry points called the same `tryRevisionDiff` planner, which always
used Unicode-scalar diffing. An English word-mode request consequently became
several edits inside words. The Force Line Diff setting also did not select a
sentence plan on this path.

The native planner now receives and uses the requested granularity. Protected
fields, equations and links remain anchors in every plan; CJK and whitespace
cleanup retain character-level edits. Full-document preflight uses the same
granularity as the later write pass, so it checks the ranges that will actually
be used. Shared draft-check callers of `textEditHunks` retain their existing
scalar default.

Three synthetic English rewrites give these plan sizes:

| Rewrite | Prior scalar hunks | Word hunks |
| --- | ---: | ---: |
| “The investigation demonstrated substantial improvement in reproducibility.” → “The analysis showed considerable gains in reliability.” | 16 | 5 |
| “The results are significant.” → “The results are substantial.” | 2 | 1 |
| “We investigated the mechanisms responsible for the observed variation.” → “We examined the processes underlying the observed differences.” | 13 | 4 |

These counts measure edit hunks, not milliseconds or Word revision records.
A fixture executing the actual revision-aware Word adapter also verifies fewer
mutation calls for the first rewrite, correct final text and unchanged earlier
deletions.

### Search-length pieces became independent mutations

The native planner splits long target text into pieces of at most 200 UTF-16
units for Word search. Previously every piece was also a separate delete or
replacement command. Search's length limit need not dictate write granularity.

Pieces from a single editable segment are now combined into one native write
range. Before any write, a batched independent native/OOXML read verifies the
combined content and rejects hidden text or protected objects. Segment boundaries
at old deletions and inserted revisions remain disjoint; the full baseline
fingerprint is rechecked afterward. A 260-character plain deletion still uses
bounded searches but now issues one deletion instead of two.

Logs now report granularity, hunk count, search-target count and actual mutation
command count without including document excerpts. They distinguish location
overhead from write fragmentation in subsequent native-host tests.

## Costs deliberately retained

The section preflight and the write pass both locate and validate text. Inner
strategy read-back is followed by an outer paragraph check and a section check.
This creates native reads and round trips, including OOXML exports. Some checks
overlap, but removing them without a precise replacement could let a late
baseline failure leave earlier paragraphs partially written or could miss
changed equations and fields. This fix does not remove those safety boundaries.

There is no evidence that replacing the library wholesale, accepting existing
revisions, or clearing Undo more frequently would resolve the post-apply lag.

## Validation and remaining native-host test

Regression coverage includes strategy dispatch, old hidden deletions, unsafe
union rejection before writes, search-length limits, surrogate-safe pieces,
equation/link protection, no-op behavior and existing partial-write handling.
The suite models Word ranges and does not reproduce Word's layout engine.

For an actual performance comparison, open two copies of the same saved baseline
and apply the same staged amendment with the old and new builds. Compare native
mutation counts, apply duration and scrolling before closing either document.
Keep markup view and Word version fixed. A real improvement must be measured
before a document reload; merely showing correct final text is not proof that
the scrolling problem has been fixed.

Microsoft's [batching guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/correlated-objects-pattern)
supports reducing unnecessary round trips. The Mac-specific
[insertText performance report #6721](https://github.com/OfficeDev/office-js/issues/6721)
also shows that batching alone may not remove per-insertion cost. Neither source
proves the cause of Claric's persistent post-apply scrolling lag.
