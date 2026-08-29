# Changelog

All notable changes to Claric are documented here. The format is a loose
Keep-a-Changelog style; versions track `package.json` and the `v*` git tags
that drive the GHCR image publish in CI.

## [Unreleased]

### Fixed

- **retryFailedChunks** — "Click to retry failed chunks" could never
  succeed: retry chunks were rebuilt as text-only stubs without a
  `paragraphs` field (crashing the orchestrator's composer) and were handed
  a null document context. Retries now re-drive the original chunk objects,
  and the orchestrator treats a null context as "no context prefix" instead
  of crashing. (`word-actions.js`, `orchestrator.js`)
- **Selection-staleness guard on Apply** — applying a staged selection
  amendment re-read the live selection without verifying it still matched
  the staged text; a moved selection made granular diffing throw and the
  old fallback overwrote whatever was *currently* selected with the stale
  amendment. The apply now refuses with a card warning when the selection
  drifted (whitespace/line-ending normalization tolerated).
  (`word-actions.js`, `conversation.js`)
- **Activity-log drawer could never open** — `#logBtn` was bound twice
  (bootstrap + `initStatusBar`), so one click toggled the drawer open and
  instantly closed. (`status-bar.js`, `taskpane.js`)
- **Task-planner output contract** — the strict OUTPUT CONTRACT enum omitted
  `image_management`/`table_management` while the CAPABILITIES section
  taught them, so a contract-obeying model could never emit those task
  types and the corresponding compound sub-tasks misrouted.
  (`task-planner.js`)
- **Trailing-comma JSON cleanup corrupted string contents** — the shared
  `,\s*([}\]])` regex also fired *inside* JSON string literals, silently
  deleting characters from cell text / tool args before they reached the
  document. Replaced by a string-aware, parse-failure-only recovery.
  (`tool-loop.js`, `table-patch.js`, `table-ops.js`)
- **Truncation detection** — streams that closed without `[DONE]` and
  without `finish_reason` were returned as complete answers, and
  non-streaming responses with `finish_reason=length` flowed into the diff
  pipeline as if complete. Both now fail loudly; the SSE reader also
  flushes the decoder's final multi-byte tail and releases the body after
  `[DONE]`. (`llm-client.js`)
- **Session persistence on quota pressure** — oversized sessions lost only
  illustration previews, so a large document run could exceed the per-session
  cap and silently vanish from history. The trimmer now degrades in stages
  (previews → proposal diffs → proposals → pathological message text), and
  a failed history-index write surfaces an error instead of leaving the
  index and blobs inconsistent. The total-cap check also stopped
  re-parsing every stored session on each committed turn.
  (`sessions.js`)
- **Chunk-bookmark anchors** — `bookmarkChunkRanges` indexed paragraphs from
  a previous `Word.run` without validation; concurrent document changes
  between parse and bookmark made amendments land on the wrong paragraphs.
  Boundary paragraph texts are now verified before bookmarking, and
  leftover `_wdp*` bookmarks from runs interrupted by a reload are reaped
  at startup. (`reassembler.js`, `taskpane.js`)

### Changed

- **Apply concurrency guards** — proposal-card applies previously ran with
  the busy flags down (except the document-amendment card), so an apply
  could race a new turn's parse → bookmark pass or a second card's apply.
  Every card now registers its apply controller with app state (so
  Stop/cancel reaches it), locks the input while applying, and a
  module-level mutex refuses a second card's Apply while one is in flight.
  The document card's unconditional busy-flag reset (which could clobber a
  newer run) is ownership-checked. (`proposal-card.js`, `conversation.js`)
- **Shared JSON extraction** (`src/lib/json-utils.js`) — the five per-layer
  LLM-JSON parsers (tool loop, table patch, table creation, task planner,
  format ops) consolidated onto one implementation with balanced-candidate
  scanning, so prose containing a second brace pair no longer poisons the
  greedy first-`{`…last-`}` slice.
- **Log drawer bounds** — the activity drawer caps at 200 DOM entries,
  per-message work logs cap at 100 rows, and the dev-only `/log` POST
  probes once and disables itself after a failure instead of one 404 per
  log line in production. (`status-bar.js`, `chat-view.js`)
- **CSP** — the taskpane ships a Content-Security-Policy meta tag
  (defense-in-depth around the localStorage-held API key): scripts pinned
  to the bundle + Office.js CDN, `object-src 'none'`, `connect-src` left
  open for user-configured Custom endpoints. (`taskpane.html`)
- **Production proxy hygiene** — the LLM proxy no longer forwards
  `cookie`/`referer`/`origin` to upstreams. (`docker-server.cjs`)

### Removed

- Dead code and styles: unused `hasNonEmptySelection` export, pre-refactor
  `.log-drawer`/`.log-toggle` rules (the floating drawer keeps
  `.log-drawer-bar`), a dead bounds check in the orchestrator worker loop,
  undefined CSS tokens (`--font-size-xs`, `--bg-color` now defined in
  `:root`), stale accent fallbacks, and restored (static) proposal change
  lists now get the same styling as live cards.

## [0.5.0]

- Document-scope image & table tool sessions (operable objects at
  full-document polish), multi-table coordination, comprehensive table
  style tools, image align/link/alt-title, merge_cells, list ops and a
  systematic style prompt, sentence-diff toggle fix.

## [0.4.0]

- Table creation pipeline, mixed paragraph+table selections, comment
  granularity, rename to Claric, hardened docker server, dev endpoints
  split behind `ENABLE_DEV_ENDPOINTS`, vendored diff-match-patch.

(Older entries were tracked in release notes before this file existed.)
