# Architecture Documentation

## Overview

Claric is a Microsoft Word add-in that provides two core workflows:

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
    llm-client.js              # LLM API client (OpenAI-compatible; honors
                               #   per-provider apiPath prefixes)
    providers.js               # Provider catalog: presets for Ollama, vLLM,
                               #   DeepSeek, Zhipu GLM, Moonshot Kimi, Custom
    prompt-manager.js          # 4-category prompt CRUD, activation, composition
    comment-extractor.js       # Comment extraction, document text extraction,
                               #   tracked changes OOXML parsing, token estimation
    document-generator.js      # Summary document creation (markdown→HTML→Word)
    comment-queue.js           # Async comment queue with bookmark persistence
    comment-request.js         # Comment request data model
    structure-model.js         # Paragraph block model for diff strategies
  scripts/
    verify-word-api.js         # Word API version verification utility
  taskpane/                    # Chat-driven main UI
    taskpane.html              # Chat layout: header, message list, input bar,
                               #   settings slide-over
    taskpane.js                # Bootstrap only: module wiring, Office.onReady,
                               #   WordApi capability detection
    taskpane.css               # Chat styles (messages, proposal cards, pills)
    app-state.js               # Shared state: config, PromptManager, processing
                               #   flags, normalizeConfig
    skills.js                  # Skill registry: 6 built-ins + custom prompts
                               #   exposed as slash commands
    conversation.js            # Turn routing: skill / selection edit / doc Q&A,
                               #   concurrency guard, cancel
    word-actions.js            # Document/LLM pipelines with explicit args:
                               #   selection edits (prepare/apply split for
                               #   proposal cards), doc-scope chunked runs,
                               #   summary, chat Q&A
    ui/
      chat-view.js             # Message rendering, streaming text, progress
      input-bar.js             # Textarea, skill picker popup, send/cancel,
                               #   model pill
      welcome.js               # Welcome empty state with skill chips
      settings-view.js         # Settings slide-over + prompt management
      proposal-card.js         # Staged edit proposal card (Apply/Reject)
      status-bar.js            # Activity log drawer, comment pending bar,
                               #   connection status

tests/                         # Jest unit tests (514 tests, 22 suites)
  prompt-state.spec.js         # PromptManager CRUD, activation, summary category
  prompt-persistence.spec.js   # localStorage round-trip, migration, edge cases
  prompt-composition.spec.js   # composeMessages, composeSummaryMessages, placeholders
  comment-extractor.spec.js    # Comments, structured extraction, OOXML tracked changes
  document-generator.spec.js   # HTML building, markdown conversion, table borders
  comment-queue.spec.js        # Queue state management, bookmark naming
  llm-client.spec.js           # sendPrompt, stripThinkTags, testConnection
  skills.spec.js               # Skill registry, resolveSkill parsing, custom skills
  conversation.spec.js         # Turn routing, concurrency guard, cancel
  llm-stream.spec.js           # sendPromptStream SSE parsing and fallback

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
User selects text → types an instruction (or /copy-edit) in the chat input
  → conversation.js routes the turn (skill / selection edit / doc Q&A)
  → word-actions.js reads the selection via Word.run()
  → promptManager.composeMessages(category, selection, templateOverride) builds prompt
  → llmClient.sendPrompt(config, prompt) calls LLM
  → chat shows a proposal card (Apply as tracked changes / Reject)
  → on Apply: the word-diff layer applies the response as tracked changes
  → Word shows insertions/deletions with track changes enabled
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

### Diff Engine (`src/lib/word-diff/` — vendored from office-word-diff, Apache-2.0)

Cascading strategy for applying LLM-suggested text changes:
1. **Char Diff** — character-level edits for CJK text (project-original, `char-diff.js`)
2. **Token Map** — maps individual words to `Word.Range` objects, preserves character-level formatting
3. **Sentence Diff** — tokenizes by sentence boundaries, handles structural changes
4. **Block Replace** — complete replacement fallback

### Taskpane (`src/taskpane/`)

Chat-driven UI split into focused modules:
- `taskpane.js` — bootstrap only (module wiring, Office.onReady, WordApi detection)
- `app-state.js` — shared config/state; `normalizeConfig` field-by-field validation
- `conversation.js` — turn routing: slash skill → pipeline, free text + selection → staged edit, free text → document Q&A; concurrency guard + AbortController cancel
- `skills.js` — skill registry: six built-ins plus PromptManager prompts as custom slash commands
- `word-actions.js` — the pipelines (selection prepare/apply split, doc-scope chunked runs, summary, chat Q&A) with explicit args instead of DOM/active-prompt reads
- `ui/*` — chat view, input bar with skill picker and send/cancel morph, welcome chips, settings slide-over, proposal card, status bar
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
- `backend` — provider id (`'ollama'`, `'vllm'`, `'deepseek'`, `'glm'`, `'kimi'`, `'custom'`)
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
npm test          # 514 tests, 22 suites, ~2s
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
docker build -t claric:0.3.0 .
docker compose up -d
```

The manifest is regenerated inside the container at startup from
`HOST`/`PORT`/`PROTOCOL`; a stable add-in GUID is persisted (pin with
`ADDIN_GUID`). The server also proxies `/ollama` and `/vllm` to the
upstreams configured via `OLLAMA_PROXY_TARGET`/`VLLM_PROXY_TARGET`
(host.docker.internal from a container), keeping LLM traffic same-origin
so the add-in's default backend URLs work without mixed-content or CORS
configuration.

## Licensing

- **MIT License** — Word add-in codebase
- **Apache 2.0 License** — diff strategies vendored from `office-word-diff` (`src/lib/word-diff/`, see `LICENSE`/`NOTICE` there)
