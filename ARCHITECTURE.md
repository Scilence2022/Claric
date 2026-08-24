# Architecture Documentation

## Overview

Word AI Redliner is a Microsoft Word add-in that provides two core workflows:

1. **AI Redlining** — select text, send to LLM with a prompt, apply the response as word-level tracked changes
2. **Document Summary** — extract comments, document text, and tracked changes, send to LLM, generate a formatted Word document

The add-in runs as an Office.js taskpane, served over HTTPS via webpack dev server (development) or a static Node.js server (Docker/production).

## Project Structure

```
src/
  commands/                    # Office ribbon command entry points
    commands.js
    commands.html
  lib/                         # Core modules
    llm-client.js              # LLM API client (Ollama/vLLM, OpenAI-compatible)
    prompt-manager.js          # 4-category prompt CRUD, activation, composition
    comment-extractor.js       # Comment extraction, document text extraction,
                               #   tracked changes OOXML parsing, token estimation
    document-generator.js      # Summary document creation (markdown→HTML→Word)
    comment-queue.js           # Async comment queue with bookmark persistence
    comment-request.js         # Comment request data model
    structure-model.js         # Paragraph block model for diff strategies
  scripts/
    verify-word-api.js         # Word API version verification utility
  taskpane/                    # Main UI
    taskpane.html              # 4-tab prompt UI, settings, token estimation
    taskpane.js                # Orchestration: workflows, event handlers, state
    taskpane.css               # Styles including disabled tabs, token display

tests/                         # Jest unit tests (469 tests, 20 suites)
  prompt-state.spec.js         # PromptManager CRUD, activation, summary category
  prompt-persistence.spec.js   # localStorage round-trip, migration, edge cases
  prompt-composition.spec.js   # composeMessages, composeSummaryMessages, placeholders
  comment-extractor.spec.js    # Comments, structured extraction, OOXML tracked changes
  document-generator.spec.js   # HTML building, markdown conversion, table borders
  comment-queue.spec.js        # Queue state management, bookmark naming
  llm-client.spec.js           # sendPrompt, stripThinkTags, testConnection

scripts/
  generate-manifest.cjs        # Builds manifest.xml from template + .env
                               #   (stable GUID persistence, version sync,
                               #   XML escaping)
  docker-server.cjs            # Hardened production static file server:
                               #   /healthz, traversal + crash protection,
                               #   graceful shutdown, access logging

assets/                        # Add-in icons (16/32/80px)
```

## Runtime Flows

### Amendment/Comment Flow (AI Redlining)

```
User selects text → clicks "Review Selection"
  → taskpane.js reads selection via Word.run()
  → promptManager.composeMessages(category, selection) builds prompt
  → llmClient.sendPrompt(config, prompt) calls LLM
  → office-word-diff applies response as tracked changes
  → Word shows insertions/deletions with track changes enabled
```

### Summary Flow (Document Summary)

```
User clicks "Generate Summary"
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

## Core Components

### Prompt Manager (`src/lib/prompt-manager.js`)

Four categories: **context**, **amendment**, **comment**, **summary**. Each has independent CRUD, activation, and persistence via localStorage (`wordAI.prompts.{category}`, `wordAI.active.{category}`).

Key methods:
- `getActiveMode()` — returns `'summary'` | `'amendment'` | `'comment'` | `'both'` | `'none'`
- `composeMessages(category, selection)` — builds `[{role, content}]` for amendment/comment
- `composeSummaryMessages(comments, opts)` — builds messages with `{comments}`, `{whole document}`, `{tracked changes}` placeholder replacement

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

### Document Generator (`src/lib/document-generator.js`)

- `buildSummaryHtml(summaryText, comments, title)` — converts LLM markdown to HTML via `marked.parse()`, adds inline table border styles for Word rendering, builds annex with numbered source comments
- `createSummaryDocument(html, title, log)` — creates new Word document via `context.application.createDocument()`, inserts HTML into `newDoc.body`, opens document

### Diff Engine (`office-word-diff` npm package)

Cascading strategy for applying LLM-suggested text changes:
1. **Token Map** — maps individual words to `Word.Range` objects, preserves character-level formatting
2. **Sentence Diff** — tokenizes by sentence boundaries, handles structural changes
3. **Block Replace** — complete replacement fallback

### Taskpane (`src/taskpane/taskpane.js`)

Central orchestrator. Key responsibilities:
- Tab switching with mode-dependent disable logic (Amendment/Comment disabled in Summary mode)
- Settings auto-save on every input change (no Save button)
- Live token estimation via async cached Word API reads (debounced 300ms)
- Review button routing: amendment/comment → diff workflow, summary → summary workflow
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
(taskpane.js) -- corrupt localStorage data falls back to defaults instead
of crashing the add-in.

### Runtime (localStorage)

All user settings persist in `localStorage` under the `wordAI.config` key:
- `backend` — `'ollama'` or `'vllm'`
- `backends.{name}.url` — endpoint URL
- `backends.{name}.model` — selected model
- `docExtraction.richness` — `'plain'` | `'headings'` | `'structured'`
- `trackedChangesExtraction` — boolean

Prompts persist under `wordAI.prompts.{category}` and `wordAI.active.{category}`.

## Testing

```bash
npm test          # 469 tests, 20 suites, ~2s
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
docker build -t word-ai-redliner:0.3.0 .
docker compose up -d
```

The manifest is regenerated inside the container at startup from
`HOST`/`PORT`/`PROTOCOL`; a stable add-in GUID is persisted (pin with
`ADDIN_GUID`). The production server does not proxy LLM traffic -- set an
absolute endpoint URL in the add-in Settings.

## Licensing

- **MIT License** — Word add-in codebase
- **Apache 2.0 License** — `office-word-diff` library (npm dependency)
