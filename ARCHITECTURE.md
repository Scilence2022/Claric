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
    llm-client.js              # LLM client: OpenAI-compatible chat completions
                               #   (non-streaming sendMessages/sendPrompt + SSE
                               #   streaming with think-tag demux,
                               #   reasoning_content/reasoning/reasoning_details
                               #   support, non-SSE fallback, idle timeout) plus
                               #   the Anthropic Messages API transport for the
                               #   Claude preset (x-api-key auth, content-block
                               #   parsing, typed SSE events)
    model-capabilities.js      # Per-model thinking/effort profiles and wire
                               #   mappings (reasoning_effort, thinking.type,
                               #   output_config.effort, budget_tokens, ...),
                               #   temperature support rules, UI option lists
    providers.js               # Provider catalog: Ollama, vLLM, OpenAI,
                               #   Claude (Anthropic), DeepSeek, Zhipu GLM,
                               #   Moonshot Kimi, MiniMax (intl + CN),
                               #   中科大模型 (zhongkeyu.com), Custom
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
    mcp-client.js             # Minimal MCP client (Streamable HTTP JSON-RPC):
                               #   initialize handshake + session-id echo,
                               #   tools/list, tools/call; JSON and SSE
                               #   response bodies; injectable fetch
    mcp-tools.js              # MCP↔tool-loop bridge: name sanitization +
                               #   cross-server namespacing, JSON Schema →
                               #   example args, executor routing (text join,
                               #   image attachments, isError, truncation);
                               #   synthetic resource tools (list/read),
                               #   MCP prompts → skill package convergence
    skill-package.js          # SKILL.md skill package parser/serializer
                               #   (frontmatter name/description/category/scope
                               #   + markdown body → skill descriptor)
    skill-store.js            # Imported skill package persistence
                               #   (localStorage wordAI.skills.imported, cap 24)
    skill-limits.js           # Shared skill-package size caps
    file-attachments.js        # Chat file uploads: type detection (text/image/
                               #   docx/pdf), size/count caps, parsing (File API;
                               #   mammoth + pdf.js legacy lazily imported),
                               #   prompt context assembly, persistence metadata
    response-parser.js         # ===AMENDMENT===/===COMMENT=== splitting,
                               #   fallback classification prompt
    json-utils.js              # Shared LLM-JSON extraction: string-aware
                               #   trailing-comma cleanup, balanced-candidate
                               #   scanning, object/array extractors (one
                               #   implementation behind tool-loop / table-patch
                               #   / table-ops / task-planner / format-ops)
    format-ops.js              # NL→JSON formatting/insert ops: prompt builder,
                               #   strict allowlist parser, op describer
    illustration.js            # SVG illustration prompt/parse/sanitize
                               #   (DOMPurify SVG profile)/dimensions/position
    task-planner.js            # Compound-instruction decomposition prompt +
                               #   plan parser (7 task types + image_management/
                               #   table_management, caps)
    table-patch.js             # Coordinate patch protocol for multi-cell
                               #   table selections: prompt builder, JSON
                               #   patch parser/validator, row-op ordering
                               #   (planRowOpOrder is tableIndex-aware —
                               #   ascending table, then descending row);
                               #   merged-cell aware (shadow slots read-only)
    table-ops.js               # Table creation protocol: EN/ZH dimension
                               #   inference (empty-grid fast path), creation
                               #   prompt builder, strict spec parser/validator
    platform.js                # Office host detection; which hosts record
                               #   table row insert/delete as tracked revisions
    selection-with-comments.js # Splices comment anchors into selection OOXML
    selection-context.js       # Pure formatters: table selections → LLM-ready
                               #   markdown grids (budget-aware row/col
                               #   truncation), mixed paragraph+table blocks,
                               #   cursor-location context
    tool-registry.js           # L1 tool-calling: tool spec normalization +
                               #   ReAct-style loop system prompt (backend-
                               #   agnostic — no native function-calling)
    tool-loop.js               # L3 tool-calling: the execution loop (one JSON
                               #   call/turn + observation, step budget,
                               #   injected send/execute, abort-aware)
    table-model.js             # L2 tool-calling: table draft model + tools
                               #   (get_state/set_cell/insert_row/delete_row/
                               #   merge_cells + style tools set_table_style/
                               #   set_borders/set_cell_format/set_font/
                               #   set_header_row/set_layout/
                               #   set_column_widths); validates ops,
                               #   translates to tablePatch (incl. merges +
                               #   styleOps). Multi-table sessions: tools take
                               #   a tableIndex (default 1) and the patch
                               #   carries tableOriginals for per-table
                               #   staleness checks
    table-style.js             # Pure style-op vocabulary for the table tool
                               #   loop: color/border/alignment/font
                               #   normalization, target-region clipping,
                               #   card labels (WordApi 1.3 surface)
    image-model.js             # L2 tool-calling: image draft model + tools
                               #   (list_images/read_image/
                               #   design_illustration/replace_illustration/
                               #   delete_image/resize_image (widthPt |
                               #   heightPt | scalePct + lockAspectRatio)
                               #   /align_image /set_alt_text (description
                               #   + title) /set_image_link);
                               #   stable snapshot indexes, card items
    sanitize.js                # Shared lazy DOMPurify factory (used by
                               #   document-generator and illustration)
    vendor/
      diff-match-patch.js      # Pinned verbatim copy of diff-match-patch
                               #   1.0.5 (Apache-2.0; upstream abandoned),
                               #   typed via sibling .d.ts
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
    message-shape.js           # Canonical chat-message shape shared by both
                               #   legs of the session round-trip (id
                               #   generation, field validation, attachment
                               #   and citation metadata)
    conversation.js            # Turn routing (intent regexes + task planner),
                               #   per-pipeline turn runners, proposal staging,
                               #   concurrency guard, cancel
    word-actions.js            # Document/LLM pipelines with explicit args:
                               #   selection/append/format/table/illustration
                               #   prepare+apply pairs, gated doc-scope runs,
                               #   planner, Q&A, summary, selection watch,
                               #   reveal/locate
    agent-actions.js           # Word glue for the tool-calling stack:
                               #   prepareTableToolEdit (chained table edits →
                               #   tablePatch proposal), prepareImageToolEdit +
                               #   applyImageOps (image management ops)
    ui/
      chat-view.js             # Message rendering, streaming text, per-turn
                               #   work log, model activity (auto-scroll),
                               #   progress bar, citation pills
      input-bar.js             # Composer: textarea, skill picker popup + "+"
                               #   skills menu, macOS-style auto-apply toggle,
                               #   send/cancel, model pill, selection preview
                               #   chip (text + image thumbnails + +N badge),
                               #   file-upload button + attachment chips
                               #   (parse/remove/clear, validation via
                               #   lib/file-attachments.js)
      welcome.js               # Welcome empty state with skill chips
      settings-view.js         # Settings slide-over (General / Prompts / Skills / MCP Servers tabs)
                               #   + prompt management
      proposal-card.js         # Staged proposal card: per-change checkboxes,
                               #   locate button, image/table previews,
                               #   selective apply, terminal states
                               #   (applied/rejected/warning/error)
      diff-view.js             # Inline <del>/<ins> text diff element
                               #   (diff-match-patch semantic cleanup)
      status-bar.js            # Activity log drawer, comment pending bar,
                               #   connection status

tests/                         # Jest unit tests (~1360 tests, 62 suites)
  conversation.spec.js         # Turn routing (all intent families + compound +
                               #   ambiguous), staging, selective apply, warnings
  reassembler.spec.js          # Alignment, bookmarks, re-anchoring, blank
                               #   paragraphs, line endings, table-paragraph
                               #   guards, validation
  table-patch.spec.js          # Table patch prompt, JSON parsing/validation,
                               #   bounds enforcement, row-op ordering
  table-ops.spec.js            # Table creation: dimension inference (EN/ZH),
                               #   prompt builder, spec parsing/limits
  platform.spec.js             # Host detection, tracked-row-op support mapping
  word-actions-table.spec.js   # Table selection route: prepare/apply ordering,
                               #   desktop/web row-tracking split, stale guards;
                               #   mixed paragraph+table route: per-paragraph
                               #   text, table guards, truncation refusal
  word-actions-table-create.spec.js # Table creation route: empty-grid fast
                               #   path, constrained content generation,
                               #   insert positions, platform tracking split
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
  selection-context.spec.js   # Table→markdown formatting, truncation notes,
                               #   mixed paragraph+table interleaving,
                               #   cursor-location formatting
  tool-loop.spec.js           # Tool registry prompts + loop protocol
                               #   (recovery, step limit, abort)
  table-model.spec.js         # Table draft model: op validation, patch
                               #   translation, tool dispatch; style tools
                               #   (borders/format/font/header/layout/
                               #   column widths, styleOps in patch)
  table-style.spec.js         # Style-op vocabulary: colors, border specs,
                               #   alignment/font normalization, targets,
                               #   labels
  image-model.spec.js         # Image draft model: reshape/scale/lock,
                               #   align/link/alt-title; consume rules,
                               #   describeOps labels, normalizeImageLink
  agent-actions.spec.js       # Tool-loop Word glue: prepare/apply halves
                               #   + read_image attachments + selection focus
                               #   + 4xx image-strip retry + noOps answer
  input-bar.spec.js           # @jest-environment jsdom — selection preview
                               #   chip (text + thumbnails + +N badge), file-
                               #   attachment chips (add/remove/submit/clear,
                               #   validation logging)
  file-attachments.spec.js     # Type detection, size/count caps, text/image/
                               #   docx/pdf parsing (mammoth + pdf.js mocked),
                               #   context assembly + truncation, persist meta
  selection-images.spec.js    # readSelectionContent (text + inline pictures,
                               #   cap + totalImages), imageDataUrl mime
                               #   sniffing, debounced watchSelection
  generate-manifest.spec.js    # Manifest generation
  __mocks__/                   # Jest style mock

scripts/
  generate-manifest.cjs        # Builds manifest.xml from template + .env
                               #   (stable GUID persistence, version sync,
                               #   XML escaping)
  generate-icons.cjs           # Rasterizes assets/icon.svg to the manifest's
                               #   PNG sizes (16/32/64/80/128) via sharp;
                               #   runs before webpack in `npm run build` so
                               #   the deployed icons track the source mark
  docker-server.cjs            # Hardened production static file server:
                               #   /healthz, traversal + crash protection,
                               #   graceful shutdown, access logging,
                               #   opt-in LLM proxy (off by default)
  dev-e2e-middlewares.cjs      # Dev-only E2E/coding-agent endpoints,
                               #   registered by webpack dev server only when
                               #   ENABLE_DEV_ENDPOINTS=true

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
   context; a review intent without edit verbs ("检查选择的表格的内容") is
   also Q&A (analysis, not a rewrite); otherwise SELECTION_EDIT.
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
Style ops from the tool loop ride the same patch: region-bound ops
(shading/alignment/font on covered rows/cells) run while original
coordinates are still valid — ahead of the row ops — and table-level ops
(built-in table style + banding, borders incl. per-row borders, header
rows, layout, column widths, whole-table formats) run after the structure
settles; each op syncs separately so an unsupported op degrades to a
warning instead of failing the apply.

**Table creation:** a "insert a 3×3 table …" instruction routes to its own
pipeline (TURN_TYPE.TABLE). Explicit dimensions without content wording are
inferred by lib/table-ops.js into an empty-grid spec deterministically — no
LLM call. Content-bearing or dimensionless requests go to the model with a
strict JSON contract (rectangular plain-text matrix, size limits); when the
instruction stated dimensions, they are restated as a hard constraint and
the model's grid is rejected wholesale on mismatch. The card shows a
read-only grid preview; Apply re-validates the spec (proposals may have
round-tripped through session persistence) and inserts one native table via
`Body.insertTable` (start/end) or `Range.insertTable` on the selection
(before/after), then sets grid style, headerRowCount, and AutoFit. The same
platform split as row ops applies: the insertion is tracked only on Word
desktop, elsewhere it lands untracked with a warning.

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

### Selection as Object (images + tables)

Picture and table selections were previously invisible to the text-only
QA path — `selection.text` is empty for images, so the preview chip
didn't show, `hasSelection` was false, and the LLM never saw the
content. The fix is the **selection-as-object** pattern: a selected
picture or table enters the conversation as a live reference with a
tool list, never as raw content.

- **Selection reader**: `readSelectionContent` returns `{ text, images,
  totalImages, hasMultiCellTableRegion, tableRegion }`. `watchSelection`
  debounces selection-change events into the input-bar preview
  (text snippet + image thumbnail row + `+N` badge when the image cap
  truncates, plus a `Table R1C1 → R3C2` corner-coords badge when the
  selection covers multiple cells).
- **Image selection** routes any instruction to the IMAGE_TOOL session:
  snapshot indexes serve as handles; the task prompt lists each
  picture (`image 1: 300×200pt, alt "…"`) and marks the user's current
  selection. The new `read_image` tool is host-executed and attaches
  the picture as a multimodal image input to the next observation
  message; text-only backends fall back via an attachment-stripped
  retry. A read-only loop (no recorded ops, just a `finish` summary)
  becomes the chat answer — no proposal card.
- **Multi-cell table selection** routes any instruction to the
  TABLE_TOOL session (`TURN_TYPE.TABLE_TOOL` — the table-side
  counterpart of IMAGE_TOOL): `prepareTableToolEdit` already seeds the
  table draft model from the selection's table region (grid + style
  snapshot), exposing the `get_state` / `set_cell` / `insert_row` /
  `delete_row` / `merge_cells` tools plus the style tools
  (`set_table_style`, `set_borders`, `set_cell_format`, `set_font`,
  `set_header_row`, `set_layout`, `set_column_widths`) used for chained
  instructions. A format intent that names the table's look
  (边框/底纹/表头/…) with a table selection diverts here too — the
  paragraph format pipeline never touches tables. Per-cell/per-row/style
  cards reuse the existing tablePatch proposal + `applySelectionAmendment`
  table branch — zero new apply-side surface. Read-only outcomes (`finish`
  with a summary, no patch) render the summary as the chat answer —
  same read-only contract as the image tool session.
- **Intra-cell text selections** (single cell, `parentTableCell` is the
  non-null anchor) stay on the flat-text pipelines — the multi-cell
  router doesn't fire, so SELECTION_EDIT / QA behave as before.
- **Mixed text + image** selections carry text through the flat QA path
  with image metadata appended as a separate `--- SELECTED IMAGES ---
  - image N: W×Hpt, alt "…"` block; bytes are never injected.

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

### Document-Scope Image/Table Ops

The selection-as-object pattern also works at **document scope**: a
planner task of type `image_management` or `table_management` runs the
existing IMAGE_TOOL / TABLE_TOOL sessions against the WHOLE document
instead of a selection.

- **`image_management`** (`TURN_TYPE.DOCUMENT_IMAGE_TOOL`): snapshots
  every inline picture (document order, stable indexes) and drives
  `prepareImageToolEdit` — the same tool list as a picture selection.
- **`table_management`** (`TURN_TYPE.DOCUMENT_TABLE_TOOL`): reads EVERY
  table via `readDocumentTableRegions` (whole-table bounds per region,
  `tableIndex` = document order) and drives `prepareTableToolEdit` with
  the regions array. One tool loop covers ALL tables — the model picks
  a table with `tableIndex` on any tool call; the patch elements carry
  `tableIndex` and the apply side anchors each op to the matching
  `body.tables.items[i]`. The card prefixes items with `TN:` so the user
  sees which table each change belongs to.
- Each runs as a separate compound sub-task with its OWN proposal card —
  text amendments and image/table ops never mix inside one card.

Single-intent routing also fires these directly: plural-marked
instructions ("给所有图片都加上 alt 文字", "all images centered",
"给所有表格加边框", "every table to three-line") route to
DOCUMENT_IMAGE_TOOL / DOCUMENT_TABLE_TOOL without a planner round-trip;
instructions that also hit the format family (标题/居中/样式) go through
the planner (COMPOUND), which decomposes them.

### Illustration Flow

```
"设计插图并插入" → ILLUSTRATION turn → prepareIllustrationProposal()
  → LLM returns one self-contained SVG
  → sanitizeSvg() (DOMPurify SVG profile) → card with image preview
  → on Apply: SVG rasterized to PNG (offscreen canvas), inserted as a
    centered inline picture ≤450pt wide at document start (题图/头图), end,
    or inline at the caret (光标/此处 — anchor read at apply time)
```

### Tool Loop (agent-style table/image operations)

```
L4  existing proposal card + Apply        (UX and safety gating unchanged)
L3  lib/tool-loop.js                      (one JSON call/turn + observation,
                                           step budget, abort-aware; optional
                                           multimodal observation attachments —
                                           read_image attaches the picture as
                                           an image_url part on the next user
                                           message; text-only backends 4xx
                                           retry the turn with attachments
                                           stripped)
L2  lib/table-model.js, lib/image-model.js (draft models the tools operate on —
                                           NEVER Word directly; ops are a
                                           staged, diffable transaction; the
                                           table side covers cell text, row
                                           ops, merge_cells, and the style
                                           tools — table style/banding,
                                           borders (incl. row borders for
                                           three-line tables), cell shading/
                                           alignment, fonts, header rows,
                                           layout, column widths — validated
                                           by lib/table-style.js; the image
                                           side covers alt text (description
                                           + title), paragraph alignment
                                           (居中 via pic.paragraph), size
                                           (width/height/scale + aspect
                                           lock), and hyperlink set/clear)
L1  lib/tool-registry.js                  (tool specs + loop system prompt)
```

Backend-agnostic by design: the loop is ReAct-style over plain chat
messages (no native function-calling — the project targets arbitrary
OpenAI-compatible endpoints), so any model can drive it.

Triggers (single-shot protocols stay the default — the loop costs multiple
LLM round trips):
- Chained instructions on a selection ("…，然后…") → prepareTableToolEdit;
  non-table selections return null and fall back to the single-shot patch.
- An unparseable single-shot table patch retries once via the tool loop.
- IMAGE_TOOL turn: image-management instructions ("删除/替换/缩放/alt text on
  图片/插图", multi-image design), OR any instruction paired with an
  image-only selection (`routeTurn` short-circuits the hasSelection branch
  to IMAGE_TOOL when the selection carries images and no text). Selection
  metadata (width/height/altText) is matched onto snapshot indexes by
  first-match-wins so the task prompt can tell the model which picture the
  user is pointing at. The loop's read_image tool is host-executed: it
  fetches the picture base64 via Office.js and returns it as an
  `attachments: [{dataUrl}]` observation; `tool-loop` turns the
  attachment into an OpenAI-compatible multimodal content parts array on
  the next user message, so vision-capable models actually see the
  picture. Text-only backends that reject the array get one
  attachment-stripped retry (`_sendLoopMessages`). A read-only loop (no
  ops recorded but the model called `finish` with a `summary`)
  resolves to a `{noOps: true, answer}` proposal that the turn runner
  renders directly in chat — no proposal card.

The table loop's ops translate into the existing tablePatch shape, so the
table proposal card and applySelectionAmendment serve unchanged.

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
    + --- SELECTED TEXT --- (when a text selection exists)
    + --- SELECTED IMAGES --- (metadata reference when the selection also
      carries image(s); content reading is the read_image tool in the
      image session, not bytes injected here)
    + --- ATTACHED FILE: name --- sections (chat file uploads — extracted
      text from .txt/.md/.docx/.pdf, appended by conversation.submit to the
      routed question/instruction, capped at 200K chars)
    + --- DOCUMENT --- (full text at configured extraction richness)
  → sendPromptStream with 300s idle timeout; tokens stream into the message
  → uploaded image attachments (questionImages) instead send one multimodal
    user message (text + image_url parts) via sendMessagesStream; an HTTP 4xx
    (text-only backend) retries once without the images, mirroring the tool
    loop's _sendLoopMessages degradation
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
- `applyChunkResults()` — aligns LLM output to original paragraphs (LCS on exact matches, then greedy similarity matching at 0.4 threshold; CJK uses bigram similarity; blank paragraphs excluded) and applies keep/delete/insert ops in reverse document order as tracked changes; comments are inserted in a separate later phase with tracking off (avoids `AccessDenied` on ranges with pending revisions). Accepts an optional `signal` (cooperative pause: aborts between chunks, leaving the remaining chunks' bookmarks intact for resumable apply) and `onChunkApplied(chunkId, outcome)` (per-chunk progress for the proposal card). Returns `{ amendmentsApplied, commentsInserted, noChangeCount, errors, appliedChunkIds, interrupted }` for honest UI feedback — `interrupted: true` signals a Stop mid-apply that the card turns into a "Continue applying" resume.
- Table guard: insert/delete alignment ops whose target paragraph sits inside a table are skipped with a warning (they would add in-cell paragraphs or delete cell content, never rows); in-cell `keep` edits still diff granularly
- Re-anchoring: `_reanchorChunkRange()` narrows a drifted staged range to the contiguous window holding the staged paragraphs (`_findAnchorWindow` over stored original texts); if contiguity is lost, the chunk is **skipped** with an error entry rather than falling back to the raw range
- Truncation guard rejects LLM output under 30% of the original chunk length

### Format Ops (`src/lib/format-ops.js`)

- `buildFormatPrompt(instruction, scopeText, scope)` — asks the LLM for a JSON op array; rewrite-only instructions must yield `[]`
- `parseFormatOps(raw, log)` — strict allowlist sanitizing: font payload (bold, italic, strikeThrough, doubleStrikeThrough, superscript, subscript, allCaps, smallCaps, underline, highlightColor, name, color, size), paragraph payload (style/styleBuiltIn, alignment, lineSpacing, spaceBefore/After, leftIndent/rightIndent, firstLineIndent, listType bullet|number|none, listLevel 0-8), `insert` (`text` ≤ 2000 chars, `position: start|end`); targets via `match` substring (≤255 chars), `paragraphStyle`, or whole scope
- `describeFormatOp(op)` — human-readable label for the proposal card
- List ops apply via WordApi 1.3 list APIs (`_applyListOps` in word-actions.js): non-list paragraphs start/attach to ONE new list (`startNewList` + `attachToList`, `setLevelBullet`/`setLevelNumbering`); `none` detaches; `listLevel` alone re-nests existing list items; hosts without the API get a warning, not a failure

### Illustration (`src/lib/illustration.js`)

- `buildIllustrationPrompt()` — demands one self-contained SVG (no scripts/handlers/`foreignObject`)
- `parseIllustration()` — fence stripping, 256KB cap
- `sanitizeSvg()` — DOMPurify with SVG + SVG-filters profiles
- `svgDimensions()` / `ensureSvgDimensions()` — width/height from viewBox or 1200×800 default (needed for rasterization)
- `illustrationPositionFromInstruction()` — 光标/此处/cursor → caret (anchor read at apply time), 题图/头图/开头/top/header → document start, else end
- Applied via `_svgToPngBase64()` (offscreen canvas, 1600px wide) + `insertInlinePictureFromBase64`, centered, scaled to ≤450pt width

### Task Planner (`src/lib/task-planner.js`)

- `TASK_TYPES = ['insert','format','edit','append','table','illustration','qa','image_management','table_management']`, `MAX_TASKS = 6`, per-task instruction cap 500 chars
- `buildPlanPrompt(instruction, hasSelection)` — also used to classify ambiguous single instructions
- `parsePlan(raw, log)` — validated `Array<{type, instruction}> | null`
- `conversation.js` maps tasks to turns (`turnForTask`); `insert` tasks always run at document scope; `image_management` → document-scope image tool session, `table_management` → document-scope table tool session (both first-class planner task types so "润色全文 + 给所有图片加 alt 文字" decomposes into separately reviewable cards)

### Document Generator (`src/lib/document-generator.js`)

- `buildSummaryHtml(summaryText, comments, title)` — converts LLM markdown to HTML via `marked.parse()`, adds inline table border styles for Word rendering, builds annex with numbered source comments
- `createSummaryDocument(html, title, log)` — creates new Word document via `context.application.createDocument()`, inserts HTML into `newDoc.body`, opens document

### File Attachments (`src/lib/file-attachments.js`)

Chat input file uploads, pure logic (no DOM — the input bar owns the picker and chips):

- `detectAttachmentKind(name, mime)` — extension-first classification into text / image / docx / pdf (MIME fallback)
- `validateAttachment(file, existing)` — caps: 5 files, 2 MB per text-like file, 4.5 MB per image (mirrors read_image's 6M base64-char ceiling), 10 MB total; rejections carry user-facing messages surfaced via the activity log
- `parseAttachment(file)` — text via `file.text()`, images as base64 data URLs (chunked `btoa`, FileReader fallback for jsdom), .docx via mammoth's browser bundle, .pdf via pdf.js legacy build — both parsers **dynamically imported** so they stay out of the first-paint bundle (pdf.js worker copied to `dist/pdf.worker.min.mjs` by webpack, exempted from the asset-size gate)
- `buildAttachmentContext(attachments)` — labeled `--- ATTACHED FILE: name ---` sections appended by `conversation.submit` to the routed question/instruction, capped at 200K chars with an omission note; images are name-listed (their bytes travel as image_url parts on QA turns)
- `attachmentMeta(attachments)` — name/kind/size only for chat bubbles and session persistence; extracted text and data URLs never enter localStorage

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
- `conversation.js` — turn routing (see "Turn Routing" above) and per-pipeline turn runners; every `log()` line is teed into the message's work log; concurrency guard + AbortController cancel (every LLM-bearing turn wires `chatController`; compound turns thread one shared controller into their sub-tasks so a single cancel stops the whole chain)
- `skills.js` — skill registry: built-in skills plus PromptManager prompts as custom slash commands
- `message-shape.js` — the single definition of a persisted chat message (id generation, field validation, attachment/citation metadata). `sessions.js` normalizes through it on the way into localStorage and `ui/chat-view.js` on the way back into the DOM, so a new message field cannot be validated on one leg and dropped on the other
- `word-actions.js` — the pipelines (selection/append/format/illustration/cleanup prepare+apply pairs, gated doc-scope runs, planner, Q&A, summary) with explicit args instead of DOM/active-prompt reads
- `ui/chat-view.js` — message list, streaming body, per-turn work log (auto-collapse to "Worked for Ns · M steps"), model activity region (reasoning dimmed, per-section split, pin-to-bottom auto-scroll that disengages when the user scrolls up), progress bar with ETA, citation pills
- `ui/proposal-card.js` — staged proposals: per-change checkbox list with inline diffs and locate buttons, selective apply ("Applied X of Y"), per-item applied feedback (dimmed + status tag) driven by `markItemApplied`, pause/resume ("Continue applying") via `setPaused`, terminal applied/rejected/warning/error states, optional image preview (illustration)
- `ui/diff-view.js` — display-only `<del>/<ins>` inline diff via diff-match-patch `diff_cleanupSemantic`
- Settings auto-save on every input change, with an explicit Save button for visible confirmation; unchanged localStorage keys
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
- `backend` — provider id (`'ollama'`, `'vllm'`, `'openai'`, `'claude'`, `'deepseek'`, `'glm'`, `'kimi'`, `'minimax'`, `'minimax-cn'`, `'zhongkeyu'`, `'custom'`)
- `providers.{id}.url` — endpoint URL (same-origin proxy path by default)
- `providers.{id}.apiKey` — optional API key
- `providers.{id}.model` — model id (free-text with refreshable suggestions)
- `providers.{id}.apiPath` — OpenAI API prefix (`/v1`, or `/api/paas/v4` for GLM)
- `providers.{id}.thinkingLevel` — canonical level (`default`, `off`, `on`, `adaptive`, `always`, `none`, `minimal`, `low`, `medium`, `high`, `xhigh`, `max`), interpreted per model by `model-capabilities.js`
- `providers.{id}.temperature` — 0–2 (clamped to 0–1 for Claude; omitted where the model rejects it)

Older `backends.{...}` entries (v0.3.x) migrate into `providers` on load.
- `docExtraction.richness` — `'plain'` | `'headings'` | `'structured'`
- `trackedChangesExtraction` — boolean

Prompts persist under `wordAI.prompts.{category}` and `wordAI.active.{category}`.

## Testing

```bash
npm test          # ~1300 tests, 60 suites, ~1s
npm run lint      # ESLint 9 flat config (eslint.config.cjs)
npm run build     # webpack production build
npm run verify    # lint + test + typecheck + build (what CI runs, plus npm audit --omit=dev)
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
`ADDIN_GUID`). The server also proxies `/ollama`, `/vllm`, `/openai`,
`/claude`, `/deepseek`,
`/glm`, `/kimi`, `/minimax`, `/minimax-cn`, `/zhongkeyu` to the upstreams configured via
the corresponding `*_PROXY_TARGET` variables (host.docker.internal from a
container for local LLMs), keeping LLM traffic same-origin so the add-in's
default backend URLs work without mixed-content or CORS configuration.

## Licensing

- **MIT License** — Word add-in codebase
- **Apache 2.0 License** — diff strategies vendored from `office-word-diff` (`src/lib/word-diff/`, see `LICENSE`/`NOTICE` there)
