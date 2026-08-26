# Architecture Documentation

## Overview

Claric is a Microsoft Word add-in built around a chat-driven taskpane. Free-text
instructions are routed by intent to one of six pipelines:

1. **AI Redlining (edit)** — select text or address the whole document, send to LLM with a prompt, apply the response as word-level tracked changes
2. **Formatting (format/insert)** — LLM-planned style/font/paragraph ops, plus short structural inserts (e.g. an article title), applied as tracked changes
3. **Append** — generate new content against the document context and stage an append-to-end proposal
4. **Illustration** — LLM-designed SVG, sanitized and rasterized, inserted as an inline picture
5. **Cleanup** — deterministically scan for redundant empty paragraphs (no LLM) and delete them as tracked changes
6. **Document Summary** — extract comments, document text, and tracked changes, send to LLM, generate a formatted Word document
7. **Q&A (qa)** — answer questions about the document in chat, streaming

Compound instructions are decomposed by an LLM task planner into ordered tasks
across these pipelines (the planner has no `cleanup` task type — a planned
task whose instruction is a cleanup is intercepted and routed to the
deterministic pipeline). Every document mutation is staged as a proposal card —
nothing is written until the user applies it.

The add-in runs as an Office.js taskpane, served over HTTPS via webpack dev server (development) or a static Node.js server (Docker/production).

## Project Structure

```
src/
  commands/                    # Office ribbon command entry points
    commands.js
    commands.html
  lib/                         # Core modules
    llm-client.js              # OpenAI-compatible LLM client: non-streaming
                               #   (sendMessages/sendPrompt) + SSE streaming
                               #   (sendMessagesStream) with think-tag demux,
                               #   reasoning_content support, non-SSE fallback,
                               #   and idle-timeout (resets per chunk)
    providers.js               # Provider catalog: Ollama, vLLM, DeepSeek,
                               #   Zhipu GLM, Moonshot Kimi, MiniMax (intl + CN),
                               #   Custom
    prompt-manager.js          # 4-category prompt CRUD, activation, composition
    comment-extractor.js       # Comment extraction, document text extraction,
                               #   tracked changes OOXML parsing, token estimation
    document-generator.js      # Summary document creation (markdown→HTML→Word)
    comment-queue.js           # Async comment queue with bookmark persistence
    comment-request.js         # Comment request data model
    document-parser.js         # Office.js traversal → ParsedParagraph[] with
                               #   heading levels, list info, table membership
    document-chunker.js        # Token-budgeted chunking with heading boundaries,
                               #   table exclusion (hard boundary + merge
                               #   barrier), overlap context
    context-extractor.js       # Definitions/abbreviations/outline extraction,
                               #   per-chunk filtered context prefix
    orchestrator.js            # Worker-pool parallel LLM dispatch with
                               #   progress/ETA, cancellation, merged-mode
                               #   parsing, per-chunk token streaming
    reassembler.js             # Bookmarks chunk ranges; applies results as
                               #   tracked changes (paragraph-level LCS +
                               #   similarity alignment, blank-paragraph-safe);
                               #   re-anchors drifted staged ranges at apply time
    response-parser.js         # ===AMENDMENT===/===COMMENT=== splitting,
                               #   fallback classification prompt
    format-ops.js              # NL→JSON formatting/insert ops: prompt builder,
                               #   strict allowlist parser, op describer
    illustration.js            # SVG illustration prompt/parse/sanitize
                               #   (DOMPurify SVG profile)/dimensions/position
    task-planner.js            # Compound-instruction decomposition prompt +
                               #   plan parser (6 task types, caps)
    table-patch.js             # Coordinate patch protocol for multi-cell
                               #   table selections: prompt builder, JSON
                               #   patch parser/validator, row-op ordering
    platform.js                # Office host detection; which hosts record
                               #   table row insert/delete as tracked revisions
    selection-with-comments.js # Splices comment anchors into selection OOXML
    panel-actions.js           # Legacy frozen enums (kept for tests)
    structure-model.js         # Legacy ParagraphBlock token-map model
    word-diff/                 # Diff strategies, vendored from
                               #   office-word-diff (Apache-2.0, see
                               #   LICENSE/NOTICE) + project hardening
      index.js                 #   Facade re-exporting strategies + computeDiff
      token-map.js             #   Word→Range mapping, preserves char formatting
      sentence-diff.js         #   Sentence-granularity fallback
      block-replace.js         #   Full-range replacement (last resort)
      diff-wordmode.js         #   Word-level diff computation
      char-diff.js             #   Project-original CJK char-level strategy
  scripts/
    verify-word-api.js         # Word API version verification utility
  taskpane/                    # Chat-driven main UI
    taskpane.html              # Chat layout: header, message list, input bar,
                               #   selection preview, settings slide-over
    taskpane.js                # Bootstrap only: module wiring, Office.onReady,
                               #   WordApi capability detection, selection watch
    taskpane.css               # Chat styles (messages, proposal cards, pills,
                               #   work log, model activity, warning states)
    app-state.js               # Shared state: config, PromptManager, processing
                               #   flags, normalizeConfig, debounce
    skills.js                  # Skill registry: 6 built-ins + custom prompts
                               #   exposed as slash commands
    conversation.js            # Turn routing (intent regexes + task planner),
                               #   per-pipeline turn runners, proposal staging,
                               #   concurrency guard, cancel
    word-actions.js            # Document/LLM pipelines with explicit args:
                               #   selection/append/format/illustration
                               #   prepare+apply pairs, gated doc-scope runs,
                               #   planner, Q&A, summary, selection watch,
                               #   reveal/locate
    ui/
      chat-view.js             # Message rendering, streaming text, per-turn
                               #   work log, model activity (auto-scroll),
                               #   progress bar, citation pills
      input-bar.js             # Textarea, skill picker popup, send/cancel,
                               #   model pill, selection preview chip
      welcome.js               # Welcome empty state with skill chips
      settings-view.js         # Settings slide-over + prompt management
      proposal-card.js         # Staged proposal card: per-change checkboxes,
                               #   locate button, selective apply, terminal
                               #   states (applied/rejected/warning/error)
      diff-view.js             # Inline <del>/<ins> text diff element
                               #   (diff-match-patch semantic cleanup)
      status-bar.js            # Activity log drawer, comment pending bar,
                               #   connection status

tests/                         # Jest unit tests (798 tests, 34 suites)
  conversation.spec.js         # Turn routing (all intent families + compound +
                               #   ambiguous), staging, selective apply, warnings
  reassembler.spec.js          # Alignment, bookmarks, re-anchoring, blank
                               #   paragraphs, line endings, table-paragraph
                               #   guards, validation
  table-patch.spec.js          # Table patch prompt, JSON parsing/validation,
                               #   row-op ordering
  platform.spec.js             # Host detection, tracked-row-op support mapping
  word-actions-table.spec.js   # Table selection route: prepare/apply ordering,
                               #   desktop/web row-tracking split, stale guards;
                               #   mixed paragraph+table route: per-paragraph
                               #   text, table guards, truncation refusal
  llm-stream.spec.js           # SSE parsing, reasoning demux, fallback, abort,
                               #   idle timeout
  llm-client.spec.js           # Non-streaming client, stripping helpers
  proposal-card.spec.js        # Items/checkboxes/locate/selective apply/states
  chat-view.spec.js            # Model activity auto-scroll
  format-ops.spec.js           # Op parsing/sanitizing, insert ops, descriptions
  illustration.spec.js         # SVG parse/sanitize/dimensions/position
  task-planner.spec.js         # Plan parsing, caps, prompt contract
  orchestrator.spec.js         # Parallel dispatch, cancellation, streaming
  document-parser.spec.js      # Heading detection, paragraph extraction
  document-chunker.spec.js     # Chunking constraints, table exclusion/barrier
  context-extractor.spec.js    # Definitions/abbreviations/outline
  response-parser.spec.js      # Delimited response parsing
  prompt-state.spec.js         # PromptManager CRUD, activation, summary category
  prompt-persistence.spec.js   # localStorage round-trip, migration
  prompt-composition.spec.js   # composeMessages, placeholders, output rules
  comment-extractor.spec.js    # Comments, structured extraction, OOXML revisions
  comments-on-range.spec.js    # extractCommentsOnRange
  comment-queue.spec.js        # Queue state, bookmark naming
  document-generator.spec.js   # HTML building, markdown conversion, sanitizing
  config-persistence.spec.js   # normalizeConfig validation/migration
  providers.spec.js            # Provider preset catalog
  skills.spec.js               # Skill registry, resolveSkill
  char-diff.spec.js            # CJK char-level diff
  word-diff.spec.js            # Word/sentence diff modes
  selection-with-comments.spec.js # Comment anchor splicing
  generate-manifest.spec.js    # Manifest generation
  panel-actions.spec.js        # Legacy enums
  __mocks__/                   # Jest style mock

scripts/
  generate-manifest.cjs        # Builds manifest.xml from template + .env
                               #   (stable GUID persistence, version sync,
                               #   XML escaping)
  docker-server.cjs            # Hardened production static file server:
                               #   /healthz, traversal + crash protection,
                               #   graceful shutdown, access logging, LLM proxy

assets/                        # Add-in icons (16/32/64/80px)
```

## Turn Routing

`routeTurn(text, {hasSelection, skills, allowCompound})` in `conversation.js` is
a pure function, evaluated in this exact order:

1. Empty input → no turn.
2. `/skill args` match → the skill's pipeline (chat/context → Q&A, summary →
   summary doc, comment → selection or document comment, amendment → selection
   edit or document edit).
3. `countIntentFamilies(text) >= 2` → COMPOUND: the task planner
   (`planDocumentTasks`) decomposes the instruction into ordered tasks
   (`insert` / `format` / `edit` / `append` / `illustration` / `qa`), and each
   task is dispatched to its pipeline in user-stated order. Planning failure
   re-routes with compound disabled.
4. Illustration intent (`ILLUSTRATION_INTENT_RE`) → ILLUSTRATION.
5. Append intent (`APPEND_INTENT_RE`) → DOC_APPEND.
6. Format intent (`FORMAT_INTENT_RE`) → FORMAT (selection scope when a
   selection exists, else document scope).
7. Cleanup intent (`CLEANUP_INTENT_RE`, "删除多余的空段落" / "delete empty
   paragraphs") → CLEANUP — always document scope, deterministic (no LLM:
   the parser never sees blank paragraphs, so the text pipelines structurally
   cannot serve this).
8. Selection present → question lead means Q&A with the selection as focused
   context; otherwise SELECTION_EDIT.
9. Edit intent (`EDIT_INTENT_RE`, EN + ZH verbs; update/enrich verbs require a
   document-ish object) → DOC_EDIT (whole-document amendment).
10. Question lead (`QUESTION_LEAD_RE`) → DOC_QA (never planner-mediated).
11. Zero intent hits and not a question → COMPOUND, where the planner
    classifies the ambiguous instruction; falls back to DOC_QA on failure.

Planned tasks are mapped to turns by `turnForTask`; a planned `edit` task
whose instruction matches the cleanup intent is intercepted there and routed
to CLEANUP instead of the text pipelines.

## Runtime Flows

### Amendment/Comment Flow (AI Redlining)

```
User selects text → types an instruction (or /copy-edit) in the chat input
  → conversation.js routes the turn (SELECTION_EDIT)
  → word-actions.js reads the selection via Word.run()
  → promptManager.composeMessages(category, selection, templateOverride) builds prompt
  → llmClient streams the LLM response (tokens shown in Model activity)
  → chat shows a proposal card with an inline diff of the change
  → on Apply: the word-diff layer applies the response as tracked changes
  → Word shows insertions/deletions with track changes enabled
```

**Table selections:** when the selection spans multiple cells, the flat-text
pipeline cannot represent cell boundaries, so word-actions.js switches to the
table patch protocol instead: cells are sent as a coordinate grid (R1C1 …),
the LLM returns a JSON delta (`{"cells":[{row,col,text}],"rowOps":[…]}`), and
the card reviews one item per changed cell / row op. On Apply, cell text is
revised per cell with the granular diff strategies (tracked natively), then
row ops (`TableRow.insertRows`/`delete`) run in descending row order. Only
Word desktop records row insert/delete as tracked revisions — on Word for
the web the structure phase runs untracked with a warning (lib/platform.js).

**Mixed selections** (paragraphs plus table content, e.g. caption + table +
note — the selection overlaps a table without being inside it) take a third
route: readMixedTableSelection detects paragraphs whose parentTable is
non-null and switches the flat flow to paragraph granularity. The prompt
sends one line per paragraph/cell with an explicit "never merge, split,
reorder, or drop lines" rule; on Apply, the reassembler's paragraph alignment
maps lines back onto paragraphs with table guards (delete/insert ops never
touch table paragraphs, in-cell edits diff the paragraph content range).
There is deliberately no whole-selection replacement fallback — on a mixed
selection that fallback destroys the table, so alignment failures (e.g.
truncated output) surface on the card instead.

Document-scope edits (skill or free-text edit intent without selection) run the
chunked pipeline (parse → context → chunk → parallel dispatch) but are **gated**:
results are staged as a "Proposed edits to N section(s)" card with one checkbox
per amended chunk. Bookmarks persist until Apply (tracked changes, selective by
chunk) or Reject (discard + cleanup). Staged ranges are re-anchored at apply
time, so inserting content via another card first (e.g. a title) never causes
the amendment to delete drifted paragraphs.

Tables never enter this pipeline: the chunker skips inTable paragraphs and
splits chunks at table boundaries (with a merge barrier, so no chunk's
bookmark range can span a table). Flattened cell lines would invite the model
to reorganize or echo them, and the paragraph alignment would then pollute
the document with phantom paragraphs. Table content stays untouched by
document checks; use the selection routes (table patch / mixed) to edit it.

### Format Flow

```
"make all headings centered" / "增加文章标题"
  → FORMAT turn → prepareFormatProposal()
  → LLM returns a JSON op array (font/paragraph/insert ops, strict allowlist)
  → proposal card lists one checkbox per op (with locate link for match ops)
  → on Apply: ops set Word.js font/paragraph properties or insert styled
    paragraphs, recorded as tracked (Formatted) revisions
```

### Append Flow

```
"继续写..." → DOC_APPEND turn → prepareDocumentAppend()
  → LLM drafts content against the full document (+ selection as focus)
  → "Proposed content to append at the document end" card
  → on Apply: paragraphs inserted at body end as tracked changes
```

### Illustration Flow

```
"设计插图并插入" → ILLUSTRATION turn → prepareIllustrationProposal()
  → LLM returns one self-contained SVG
  → sanitizeSvg() (DOMPurify SVG profile) → card with image preview
  → on Apply: SVG rasterized to PNG (offscreen canvas), inserted as a
    centered inline picture ≤450pt wide at document start (题图/头图) or end
```

### Cleanup Flow

```
"删除多余的空段落" → CLEANUP turn → prepareEmptyParagraphCleanup()
  → Word.run scan of body.paragraphs: whitespace-only text, excluding the
    final paragraph, table-cell paragraphs, and paragraphs holding inline
    pictures → "Delete N empty paragraph(s)" proposal card
  → on Apply: re-scan at apply time and delete the empty paragraphs
    (reverse order, tracked per config); zero deletions → honest warning
```

### Summary Flow (Document Summary)

```
User runs "/summarize-contract"
  → extractAllComments() gets all document comments (WordApi 1.4)
  → extractDocumentStructured({ richness }) gets document text (if {whole document} in prompt)
  → extractTrackedChanges() parses OOXML for revisions (if {tracked changes} in prompt)
  → promptManager.composeSummaryMessages(comments, opts) builds prompt
  → llmClient.sendPrompt(config, prompt) calls LLM
  → marked.parse(response) converts markdown to HTML
  → buildSummaryHtml() adds title, summary, annex with source comments
  → createSummaryDocument() creates new Word doc via Application.createDocument()
  → New document opens with formatted content
```

### Q&A Flow

```
Question → DOC_QA → answerQuestion()
  → prompt = context prompt + optional chat-skill template + question
    + --- SELECTED TEXT --- (when a selection exists)
    + --- DOCUMENT --- (full text at configured extraction richness)
  → sendPromptStream with 300s idle timeout; tokens stream into the message
```

## Core Components

### Prompt Manager (`src/lib/prompt-manager.js`)

Four categories: **context**, **amendment**, **comment**, **summary**. Each has independent CRUD, activation, and persistence via localStorage (`wordAI.prompts.{category}`, `wordAI.active.{category}`).

Key methods:
- `getActiveMode()` — returns `'summary'` | `'amendment'` | `'comment'` | `'both'` | `'none'`
- `composeMessages(category, selection)` — builds `[{role, content}]` for amendment/comment
- `composeSummaryMessages(comments, opts)` — builds messages with `{comments}`, `{whole document}`, `{tracked changes}` placeholder replacement

### LLM Client (`src/lib/llm-client.js`)

- `sendMessages(config, messages, log, signal, timeoutMs)` — non-streaming, fixed total timeout (default 120s), WebView2-safe manual AbortController bridging
- `sendMessagesStream(config, messages, {onContent, onReasoning}, log, signal, timeoutMs)` — SSE streaming; parses `content` and `reasoning_content` deltas; `createStreamDemux` splits inline `<think>` blocks across token boundaries; falls back to non-SSE JSON delivered as one delta; **idle timeout** re-arms on response headers and every chunk, so only a stalled stream trips `TimeoutError`; returns `{content, reasoning}`
- `sendPrompt` / `sendPromptStream` — single-string wrappers
- Helpers: `stripThinkTags`, `stripMarkdown`, `stripChunkDelimiters`, `testConnection`

### Comment Extractor (`src/lib/comment-extractor.js`)

Three extraction functions:
- `extractAllComments()` — Word API `body.getComments()` with three-sync batch loading
- `extractDocumentStructured({ richness })` — paragraph-level extraction with 3 richness levels (plain/headings/structured)
- `extractTrackedChanges()` — OOXML parsing via `body.getOoxml()` + browser DOMParser

OOXML tracked changes pipeline:
1. Parse XML, extract `w:body` from `pkg:package` wrapper
2. Remove `w:proofErr` elements (normalization)
3. Process `w:del` elements, pair with adjacent `w:ins` (same author) as replacements
4. Process unpaired `w:ins` as additions
5. Process `w:moveFrom`/`w:moveTo` as move operations
6. Skip `w:ins`/`w:del` inside `w:trPr` (table row markers)
7. Extract paragraph context for each change

### Reassembler (`src/lib/reassembler.js`)

- `bookmarkChunkRanges()` — hidden `_wdp…` bookmarks persist chunk ranges across LLM processing time
- `applyChunkResults()` — aligns LLM output to original paragraphs (LCS on exact matches, then greedy similarity matching at 0.4 threshold; CJK uses bigram similarity; blank paragraphs excluded) and applies keep/delete/insert ops in reverse document order as tracked changes; comments are inserted in a separate later phase with tracking off (avoids `AccessDenied` on ranges with pending revisions)
- Table guard: insert/delete alignment ops whose target paragraph sits inside a table are skipped with a warning (they would add in-cell paragraphs or delete cell content, never rows); in-cell `keep` edits still diff granularly
- Re-anchoring: `_reanchorChunkRange()` narrows a drifted staged range to the contiguous window holding the staged paragraphs (`_findAnchorWindow` over stored original texts); if contiguity is lost, the chunk is **skipped** with an error entry rather than falling back to the raw range
- Truncation guard rejects LLM output under 30% of the original chunk length
- Returns `{ amendmentsApplied, commentsInserted, noChangeCount, errors }` for honest UI feedback

### Format Ops (`src/lib/format-ops.js`)

- `buildFormatPrompt(instruction, scopeText, scope)` — asks the LLM for a JSON op array; rewrite-only instructions must yield `[]`
- `parseFormatOps(raw, log)` — strict allowlist sanitizing: font payload (bold, italic, strikeThrough, doubleStrikeThrough, superscript, subscript, allCaps, smallCaps, underline, highlightColor, name, color, size), paragraph payload (style/styleBuiltIn, alignment, lineSpacing, spaceBefore/After, leftIndent/rightIndent, firstLineIndent), `insert` (`text` ≤ 2000 chars, `position: start|end`); targets via `match` substring (≤255 chars), `paragraphStyle`, or whole scope
- `describeFormatOp(op)` — human-readable label for the proposal card

### Illustration (`src/lib/illustration.js`)

- `buildIllustrationPrompt()` — demands one self-contained SVG (no scripts/handlers/`foreignObject`)
- `parseIllustration()` — fence stripping, 256KB cap
- `sanitizeSvg()` — DOMPurify with SVG + SVG-filters profiles
- `svgDimensions()` / `ensureSvgDimensions()` — width/height from viewBox or 1200×800 default (needed for rasterization)
- `illustrationPositionFromInstruction()` — 题图/头图/开头/top/header → document start, else end
- Applied via `_svgToPngBase64()` (offscreen canvas, 1600px wide) + `insertInlinePictureFromBase64`, centered, scaled to ≤450pt width

### Task Planner (`src/lib/task-planner.js`)

- `TASK_TYPES = ['insert','format','edit','append','illustration','qa']`, `MAX_TASKS = 6`, per-task instruction cap 500 chars
- `buildPlanPrompt(instruction, hasSelection)` — also used to classify ambiguous single instructions
- `parsePlan(raw, log)` — validated `Array<{type, instruction}> | null`
- `conversation.js` maps tasks to turns (`turnForTask`); `insert` tasks always run at document scope

### Document Generator (`src/lib/document-generator.js`)

- `buildSummaryHtml(summaryText, comments, title)` — converts LLM markdown to HTML via `marked.parse()`, adds inline table border styles for Word rendering, builds annex with numbered source comments
- `createSummaryDocument(html, title, log)` — creates new Word document via `context.application.createDocument()`, inserts HTML into `newDoc.body`, opens document

### Diff Engine (`src/lib/word-diff/` — vendored from office-word-diff, Apache-2.0)

Cascading strategy for applying LLM-suggested text changes:
1. **Char Diff** — character-level edits for CJK text (project-original, `char-diff.js`)
2. **Token Map** — maps individual words to `Word.Range` objects, preserves character-level formatting
3. **Sentence Diff** — tokenizes by sentence boundaries, handles structural changes
4. **Block Replace** — complete replacement fallback

Local modifications on top of upstream (kept intentionally small):
- Every strategy accepts `options.trackChanges` (default `true`); when `false`
  it does not touch the document's `changeTrackingMode` — the caller owns it.
  Callers always set the mode explicitly (on or off) from
  `config.trackChangesEnabled`, so the settings toggle is honored end to end,
  and a mid-edit failure can never leak `trackAll` (finally-restore).
- Sentence-mode diffs the occurrence-ordered sentence sequence (upstream
  diffed the deduped list, silently misaligning repeated sentences).
- Token-map resolves the Nth occurrence of a token to the Nth search match
  inside a coarse range (upstream always took the first match).
- Fallback resets run with tracking off so they don't add spurious revisions.

### Taskpane (`src/taskpane/`)

Chat-driven UI split into focused modules:
- `taskpane.js` — bootstrap only (module wiring, Office.onReady, WordApi detection, selection watch wiring)
- `app-state.js` — shared config/state; `normalizeConfig` field-by-field validation
- `conversation.js` — turn routing (see "Turn Routing" above) and per-pipeline turn runners; every `log()` line is teed into the message's work log; concurrency guard + AbortController cancel
- `skills.js` — skill registry: six built-ins plus PromptManager prompts as custom slash commands
- `word-actions.js` — the pipelines (selection/append/format/illustration/cleanup prepare+apply pairs, gated doc-scope runs, planner, Q&A, summary) with explicit args instead of DOM/active-prompt reads
- `ui/chat-view.js` — message list, streaming body, per-turn work log (auto-collapse to "Worked for Ns · M steps"), model activity region (reasoning dimmed, per-section split, pin-to-bottom auto-scroll that disengages when the user scrolls up), progress bar with ETA, citation pills
- `ui/proposal-card.js` — staged proposals: per-change checkbox list with inline diffs and locate buttons, selective apply ("Applied X of Y"), terminal applied/rejected/warning/error states, optional image preview (illustration)
- `ui/diff-view.js` — display-only `<del>/<ins>` inline diff via diff-match-patch `diff_cleanupSemantic`
- Settings auto-save on every input change (no Save button), unchanged localStorage keys
- WordApi version detection and feature gating (1.4 for comments)

## Configuration

### Build-Time (.env → manifest.xml)

`scripts/generate-manifest.cjs` reads `.env` and generates `manifest.xml` from `manifest.template.xml`. Runs automatically from webpack config and at container startup.

```
HOST=localhost       # Hostname reachable from Word
PORT=3000           # Port for HTTPS server
PROTOCOL=https      # Must be https for Office add-ins
ADDIN_GUID=...      # Optional: pin a stable add-in identity
```

The manifest `<Version>` is synced from `package.json`; the `<Id>` GUID is
generated on first run and persisted to `.manifest-guid` so the add-in
identity survives restarts. All substituted values are XML-escaped.

Persisted settings are validated field-by-field by `normalizeConfig`
(src/taskpane/app-state.js) -- corrupt localStorage data falls back to defaults instead
of crashing the add-in.

### Runtime (localStorage)

All user settings persist in `localStorage` under the `wordAI.config` key:
- `backend` — provider id (`'ollama'`, `'vllm'`, `'deepseek'`, `'glm'`, `'kimi'`, `'minimax'`, `'minimax-cn'`, `'custom'`)
- `providers.{id}.url` — endpoint URL (same-origin proxy path by default)
- `providers.{id}.apiKey` — optional API key
- `providers.{id}.model` — model id (free-text with refreshable suggestions)
- `providers.{id}.apiPath` — OpenAI API prefix (`/v1`, or `/api/paas/v4` for GLM)

Older `backends.{...}` entries (v0.3.x) migrate into `providers` on load.
- `docExtraction.richness` — `'plain'` | `'headings'` | `'structured'`
- `trackedChangesExtraction` — boolean

Prompts persist under `wordAI.prompts.{category}` and `wordAI.active.{category}`.

## Testing

```bash
npm test          # 708 tests, 29 suites, ~1s
npm run lint      # ESLint 9 flat config (eslint.config.cjs)
npm run build     # webpack production build
npm run verify    # lint + test + build (what CI runs)
```

Tests run in node or jsdom environments (per-spec `@jest-environment`
docblock) with mocked Word API globals. TDD workflow: failing tests written
before implementation for each feature.

## Docker

Three-stage build on Node 22 Alpine: (1) builder compiles webpack, (2) a
deps stage installs production dependencies only, (3) the runtime stage
runs as the non-root `node` user and serves static files via
`scripts/docker-server.cjs` (hardened: traversal/malformed-URL rejection,
405 for non-GET, `/healthz` endpoint, graceful SIGTERM shutdown).

```bash
docker build -t claric .
docker compose up -d
```

The manifest is regenerated inside the container at startup from
`HOST`/`PORT`/`PROTOCOL`; a stable add-in GUID is persisted (pin with
`ADDIN_GUID`). The server also proxies `/ollama`, `/vllm`, `/deepseek`,
`/glm`, `/kimi`, `/minimax`, `/minimax-cn` to the upstreams configured via
the corresponding `*_PROXY_TARGET` variables (host.docker.internal from a
container for local LLMs), keeping LLM traffic same-origin so the add-in's
default backend URLs work without mixed-content or CORS configuration.

## Licensing

- **MIT License** — Word add-in codebase
- **Apache 2.0 License** — diff strategies vendored from `office-word-diff` (`src/lib/word-diff/`, see `LICENSE`/`NOTICE` there)
