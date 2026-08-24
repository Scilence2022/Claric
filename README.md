# Word AI Redliner

AI-powered Microsoft Word add-in that applies word-level tracked changes using
a structure-aware diff strategy — plus document summarization with comment
extraction and tracked changes analysis.

<p align="center">
  <a href="https://www.youtube.com/watch?v=SusffH8eT-Y">
    <img src="docs/word-ai-redliner.gif" alt="Word AI Redliner demo" />
  </a>
</p>

**Project history**: This library was extracted from a private codebase and open-sourced as a standalone project in Jan 2026.

## Features

### Core: AI Redlining (v0.1.0)
- Word-level diffs with tracked changes via [office-word-diff](https://github.com/niclasgrunworked/office-word-diff)
- Token map strategy with sentence fallback, block replace as last resort
- Configurable LLM backends: Ollama and vLLM (OpenAI-compatible)

### v0.2.0: Prompt System + Document Summary

**Three-Category Prompt System**
- Context, Amendment, and Comment prompt categories with dedicated tabs
- Full CRUD: create, save, update, delete prompts per category
- Per-category activation with `{selection}` placeholder replacement
- Prompts persist in localStorage across sessions

**Document Comment Summary**
- 4th "Summary" tab — extract all document comments, send to LLM, generate a formatted Word document
- `{comments}` placeholder inserts structured comment data (author, annotated text, comment text)
- `{whole document}` placeholder extracts full document text with configurable richness:
  - **Plain** — raw paragraph text
  - **Headings** — markdown-style heading markers (`## Section Title`)
  - **Structured** — headings + list item numbering and indentation
- `{tracked changes}` placeholder extracts revision marks via OOXML parsing (w:ins, w:del, w:moveFrom, w:moveTo)
- Generated summary document includes annex with numbered source comments
- LLM markdown output auto-converted to HTML via [marked](https://github.com/markedjs/marked)
- Tables in generated documents render with visible borders

**Tracked Changes Extraction (OOXML)**
- Parses `body.getOoxml()` with browser DOMParser — no external dependencies
- Handles `pkg:package` wrapper, `w:proofErr` normalization
- Pairs adjacent `w:del` + `w:ins` from same author as replacements
- Detects move operations (`w:moveFrom` / `w:moveTo`)
- Skips table row revision markers (`w:ins`/`w:del` inside `w:trPr`)
- Namespace-aware querying with prefix fallback for cross-browser compatibility
- Author identity prominently included in LLM-formatted output

**Async Comment Queue**
- Bookmark-based range persistence for async comment insertion
- Comment status bar with pending count and retry-on-error
- WordApi 1.4 detection with graceful degradation

**Settings & UX**
- Settings auto-save on every change (no Save button)
- Live token estimation with real document metrics (async Word API read, cached, debounced)
- Document extraction richness dropdown (Summary mode only)
- Tracked changes extraction toggle (Summary mode only)
- Mode switching: Amendment/Comment tabs disabled when Summary is active
- Review button relabels to "Generate Summary" in Summary mode

**Backend Selector**
- Providers: Ollama, vLLM, DeepSeek, Zhipu GLM, Moonshot Kimi, and Custom (any OpenAI-compatible endpoint)
- Unified OpenAI-compatible chat API; per-provider API prefix handled automatically (GLM uses `/api/paas/v4`)
- Cloud providers proxied same-origin (`/deepseek`, `/glm`, `/kimi`) by the dev and production servers -- no CORS setup, API key entered in Settings
- Model field is typeable with a refreshable suggestion list (Refresh button re-queries the provider's models endpoint)
- Configurable endpoint URL and optional API key per provider
- Track Changes and Line Diff toggles

### v0.3.0: Whole-Document Processing

**Parallel LLM Orchestration**
- Full-document amendment and commenting: parse, chunk, dispatch to LLM in parallel, reassemble
- Worker-pool concurrency with configurable limits (auto-tuned: 4 workers for large chunks, 6 for smaller)
- AbortController-based cancellation stops pending chunks immediately
- Progress tracking with per-chunk ETA estimation
- Retry failed chunks without re-processing successful ones

**Document Parsing & Chunking**
- Paragraph-level document parsing with style and heading detection
- Token-aware chunking with configurable max size (default 6K tokens)
- Heading-based chunk boundaries (H1/H2 trigger splits; H3+ stay coherent)
- Overlap paragraphs provide preceding context to each chunk
- Tiny trailing chunks merged into previous chunk to prevent orphans

**Context Extraction**
- Automatic definition extraction via regex (means/shall-mean/is-defined-as, the-X, hereinafter-X patterns)
- Abbreviation expansion via word-initial matching heuristic
- Document outline generation from heading hierarchy
- Context prefix formatted and injected into each chunk's LLM system message

**Formatting-Preserving Reassembly**
- Paragraph-level amendment strategy: aligns LLM output paragraphs to original document paragraphs using LCS + word-level similarity matching
- Within-paragraph word-level diff via token map strategy preserves run-level formatting (bold, italic, font, color)
- Paragraph properties (styles, numbering, indentation) preserved through paragraph-scoped operations
- Graceful degradation chain: paragraph-level -> word-level diff -> sentence diff -> block replace
- Line ending normalization (`\r` <-> `\n`) throughout the pipeline
- Content validation rejects severely truncated LLM output before applying
- Amendments applied in reverse document order to prevent range invalidation
- Bookmarks persist chunk ranges across LLM processing time

**Merged Amendment + Comment Mode**
- Comment instructions persisted with prompt data (save/restore across sessions)
- When comment instructions are provided in amendment mode, LLM produces delimited `===AMENDMENT===` / `===COMMENT===` output
- Response parser extracts both sections; comments inserted on bookmarked ranges after all amendments
- Fallback: undelimited responses treated as amendment-only

**LLM Output Quality**
- Critical output rules appended to amendment prompts: no commentary, no markdown, preserve structure
- `stripMarkdown()` post-processor as safety net for amendment responses
- `stripThinkTags()` removes `<think>` reasoning blocks from LLM output

**Testing**
- 469 unit tests across 20 test suites (Jest)
- TDD workflow: failing tests written before implementation
- Covers: prompt state/persistence/composition, comment extraction, document generation, tracked changes OOXML parsing, orchestrator dispatch/concurrency, reassembler paragraph alignment, document chunking, context extraction

## Setup

There are **two ways** to run this add-in:

| Method | Best for | Requirements |
|--------|----------|--------------|
| **Docker** | Quick setup, no Node.js needed | Docker, Docker Compose |
| **npm** | Development, customization | Node.js 22+ |

Both methods require HTTPS certificates trusted by the machine running Word.

---

## Option A: Docker (Recommended for Quick Setup)

### Prerequisites

- Docker and Docker Compose
- HTTPS certificate files (see [Create HTTPS Certificates](#create-https-certificates))

### Step-by-Step

1. **Clone the repository**

```bash
git clone https://github.com/Scilence2022/word-ai-redliner.git
cd word-ai-redliner
```

2. **Create HTTPS certificates** (see [Create HTTPS Certificates](#create-https-certificates))

   Place `server.pem` and `server-key.pem` in the project root.

3. **Configure environment variables**

   Copy the Docker example and edit it:

```bash
cp .env.docker.example .env
```

   On Windows PowerShell:

```powershell
Copy-Item .env.docker.example .env
```

   **Important:** Edit `.env` and set `HOST` to the hostname or IP address that
   the Word client can reach. If Word runs on a different machine, do **not**
   use `localhost`.

4. **Start the container**

```bash
docker compose up -d
```

   The container runs as a non-root user, serves `dist/` over HTTPS, and
   regenerates `manifest.xml` from your `.env` values on every startup. A
   stable add-in GUID is generated on first start and persisted inside the
   container (pin it with `ADDIN_GUID` in `.env` to survive container
   recreation).

5. **Get the manifest**

   Download it from the running server: `https://<HOST>:<HOST_PORT>/manifest.xml`
   (e.g. open that URL in a browser and save the file).

6. **Trust the certificate on Windows** (see [Trust the Certificate on Windows](#trust-the-certificate-on-windows))

7. **Sideload the add-in** (see [Sideload the Add-in](#sideload-the-add-in))

8. **Point the add-in at your LLM backend**

   The container proxies the `/ollama` and `/vllm` paths to the upstreams in
   your `.env` (defaults: `host.docker.internal:11434` / `:8026`), so the
   add-in's default Settings work with no extra configuration -- the LLM
   traffic stays same-origin, avoiding mixed-content and CORS issues. For a
   remote backend, set the `OLLAMA_PROXY_TARGET`/`VLLM_PROXY_TARGET` values
   (or enter an absolute URL in the add-in Settings).

---

## Option B: npm (For Development)

### Prerequisites

- Node.js 22+ (see `.nvmrc`; `engines` field enforces this)
- HTTPS certificate files (see [Create HTTPS Certificates](#create-https-certificates))

### Step-by-Step

1. **Clone the repository**

```bash
git clone https://github.com/Scilence2022/word-ai-redliner.git
cd word-ai-redliner
```

2. **Install dependencies**

```bash
npm install
```

3. **Create HTTPS certificates** (see [Create HTTPS Certificates](#create-https-certificates))

   Place `server.pem` and `server-key.pem` in the project root.

4. **Configure environment variables**

   Copy the example and edit it:

```bash
cp .env.example .env
```

   On Windows PowerShell:

```powershell
Copy-Item .env.example .env
```

   **Important:** Edit `.env` and set `HOST` to the hostname or IP address that
   the Word client can reach. If Word runs on a different machine, do **not**
   use `localhost`.

5. **Start the dev server**

```bash
npm start
```

   This generates `manifest.xml` from your `.env` values and starts the webpack
   dev server with hot reload.

6. **Trust the certificate on Windows** (see [Trust the Certificate on Windows](#trust-the-certificate-on-windows))

7. **Sideload the add-in** (see [Sideload the Add-in](#sideload-the-add-in))

   Use the `manifest.xml` file in the project root.

---

## Create HTTPS Certificates

The add-in must be served over HTTPS. Word will block untrusted certificates.

Place your cert files in the project root:

- `server.pem` (certificate)
- `server-key.pem` (private key)

### Option 1: mkcert (Recommended)

1. Install [mkcert](https://github.com/FiloSottile/mkcert).
2. Create a local CA and generate a cert:

```bash
mkcert -install

# For localhost (same machine):
mkcert localhost

# For a remote server (use your actual IP or hostname):
mkcert <your-server-ip-or-hostname>
```

3. Rename the output files:

```bash
cp localhost.pem server.pem
cp localhost-key.pem server-key.pem
```

   On Windows PowerShell:

```powershell
Copy-Item localhost.pem server.pem
Copy-Item localhost-key.pem server-key.pem
```

### Option 2: OpenSSL (Manual)

```bash
# Replace <YOUR_HOST> with localhost or your server IP/hostname
openssl req -x509 -nodes -days 365 \
  -newkey rsa:2048 \
  -keyout server-key.pem \
  -out server.pem \
  -subj "/CN=<YOUR_HOST>"
```

---

## Trust the Certificate on Windows

On the Windows PC running Word:

1. Copy the `.pem` cert file to the Windows PC.
2. Convert PEM to CRT (if needed):

```powershell
openssl x509 -in server.pem -out server.crt
```

3. Open **certmgr.msc** (run as Administrator).
4. Navigate to **Trusted Root Certification Authorities** → **Certificates**.
5. Right-click → **All Tasks** → **Import...**
6. Select the `.crt` file and finish the wizard.

**If you used mkcert**, you can install the mkcert root CA on Windows instead:

- Copy the root CA from the server machine (find it via `mkcert -CAROOT`)
- Import it into **Trusted Root Certification Authorities**

---

## Sideload the Add-in

### Word on Windows

**Method 1: Add from file**

1. Open Word → **Insert** → **Get Add-ins** → **My Add-ins**.
2. Click **Add a custom add-in** → **Add from file...**.
3. Select `manifest.xml` and confirm.

**Method 2: Network shared folder (Windows only)**

1. Create a shared folder and note the network path.
2. In Word: **File** → **Options** → **Trust Center** → **Trust Center Settings** →
   **Trusted Add-in Catalogs** → **Add catalog** (check **Show in Menu**).
3. Copy `manifest.xml` into the shared folder.
4. In Word: **Home** → **Add-ins** → **Advanced** → **Shared Folder** → select the add-in → **Add**.

For full details, see the [Microsoft sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

### Word on Mac

Mac's Word has no "Add from file" dialog -- sideloading uses the `wef`
container folder (see [Microsoft's Mac sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)):

1. Make sure the server is running (`docker compose up -d`).
2. Open Finder and press ⌘+Shift+G, then enter:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   (create the `wef` folder if it does not exist).
3. Copy the `manifest.xml` into that folder. The file lives on the server:
   `https://<HOST>:<HOST_PORT>/manifest.xml` (save it from a browser).
4. Fully quit Word (⌘Q) and reopen it -- loaded only at launch.
5. Go to the **Home** tab → **Add-ins** → select **Word AI Redliner**.
6. Trust the certificate in Keychain if prompted.

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Word shows "blocked because it isn't signed" | Trust the HTTPS certificate on the Windows client |
| Word cannot load the add-in | Verify `HOST` in `.env` is reachable from Word |
| Manifest not generated | Ensure `.env` exists before running `npm start`; with Docker, fetch it from `https://<HOST>:<HOST_PORT>/manifest.xml` |
| Firewall issues | Allow inbound TCP 3000 (or your `HOST_PORT`) on the server |
| LLM connection fails in production | Check `OLLAMA_PROXY_TARGET`/`VLLM_PROXY_TARGET` in `.env` and that the upstream is running; the container reaches the host via `host.docker.internal` |
| Container restarts repeatedly | Check `docker logs` -- missing SSL cert files or a broken build exit with a clear message |

---

## Environment Variables

### All deployments

| Variable | Default | Description |
|----------|---------|-------------|
| `HOST` | `localhost` | Hostname for manifest URLs (must be reachable from Word) |
| `PORT` | `3000` | Port for manifest URLs / server listen port |
| `PROTOCOL` | `https` | Protocol for manifest URLs and the server |
| `SSL_CERT_FILE` | `server.pem` | Path to SSL certificate |
| `SSL_KEY_FILE` | `server-key.pem` | Path to SSL private key |
| `ADDIN_GUID` | *(generated)* | Stable add-in identity; auto-generated and persisted to `.manifest-guid` on first manifest generation. Pin it in Docker so container recreation keeps the identity. |
| `OLLAMA_PROXY_PATH` | `/ollama` | Proxy path for the Ollama backend (empty disables) |
| `OLLAMA_PROXY_TARGET` | `http://localhost:11434` | Upstream Ollama base URL (`http://host.docker.internal:11434` in Docker) |
| `VLLM_PROXY_PATH` | `/vllm` | Proxy path for the vLLM backend (empty disables) |
| `VLLM_PROXY_TARGET` | `http://localhost:8026` | Upstream vLLM base URL (`http://host.docker.internal:8026` in Docker) |
| `DEEPSEEK_PROXY_PATH` | `/deepseek` | Proxy path for DeepSeek (empty disables) |
| `DEEPSEEK_PROXY_TARGET` | `https://api.deepseek.com` | Upstream DeepSeek API origin |
| `GLM_PROXY_PATH` | `/glm` | Proxy path for Zhipu GLM (empty disables) |
| `GLM_PROXY_TARGET` | `https://open.bigmodel.cn` | Upstream Zhipu GLM API origin |
| `KIMI_PROXY_PATH` | `/kimi` | Proxy path for Moonshot Kimi (empty disables) |
| `KIMI_PROXY_TARGET` | `https://api.moonshot.cn` | Upstream Moonshot Kimi API origin |
| `LLM_PROXY_TIMEOUT_MS` | `300000` | Proxy upstream timeout in ms |

### Dev server only (webpack)

| Variable | Default | Description |
|----------|---------|-------------|
| `DEV_SERVER_HOST` | `0.0.0.0` | Host to bind webpack dev server |
| `DEV_SERVER_PORT` | `3000` | Port for webpack dev server |
| `OLLAMA_PROXY_PATH` | `/ollama` | Local proxy path for LLM requests |
| `OLLAMA_PROXY_TARGET` | `http://localhost:11434` | Upstream Ollama server URL |
| `VLLM_PROXY_PATH` | `/vllm` | Local proxy path for vLLM requests |
| `VLLM_PROXY_TARGET` | `http://localhost:8026` | Upstream vLLM server URL |
| `DEFAULT_OLLAMA_URL` | `/ollama` | Default Ollama URL shown in UI |
| `DEFAULT_VLLM_URL` | `/vllm` | Default vLLM URL shown in UI |
| `DEFAULT_MODEL` | `gpt-oss:20b` | Default Ollama model shown in UI |
| `VLLM_MODEL` | `qwen3.5-35b-a3b` | Default vLLM model shown in UI |

Users can override the default URLs and models via the add-in settings UI;
those overrides persist in localStorage.

### Docker only (docker-compose)

| Variable | Default | Description |
|----------|---------|-------------|
| `HOST_PORT` | `3000` | Host-side port published by compose; also used for manifest URLs (the container listens internally on `PORT`) |

---

## Docker Image

Pre-built images are available on GitHub Container Registry:

```bash
docker pull ghcr.io/yuch85/word-ai-redliner:0.3.0
docker pull ghcr.io/yuch85/word-ai-redliner:latest
```

The image is a three-stage build on `node:22-alpine`: runtime layers contain
production dependencies only, run as the non-root `node` user, and expose a
`/healthz` endpoint used by the Docker `HEALTHCHECK`.

---

## Project Structure

See `ARCHITECTURE.md` for details.

## Testing

```bash
npm test          # 469 tests across 20 suites
npm run lint      # ESLint (flat config)
npm run build     # webpack production build
npm run verify    # lint + test + build (same as CI)
```

Test suites cover:
- `prompt-state.spec.js` — PromptManager CRUD, activation, persistence, summary category
- `prompt-persistence.spec.js` — localStorage round-trip, migration, edge cases
- `prompt-composition.spec.js` — composeMessages, composeSummaryMessages, placeholder replacement, output rules
- `comment-extractor.spec.js` — extractAllComments, extractDocumentStructured, estimateTokenCount, extractTrackedChanges (OOXML parsing)
- `document-generator.spec.js` — buildSummaryHtml (markdown conversion, table borders, escaping), createSummaryDocument (Word API)
- `comment-queue.spec.js` — CommentQueue state management, bookmark naming
- `llm-client.spec.js` — sendPrompt, stripThinkTags, stripMarkdown, testConnection
- `document-parser.spec.js` — parseDocument, paragraph extraction, heading detection, style mapping
- `document-chunker.spec.js` — chunkDocument, heading-based splitting, overlap, token limits
- `context-extractor.spec.js` — extractContext, definitions, abbreviations, outline generation
- `orchestrator.spec.js` — processChunksParallel, concurrency, cancellation, merged mode parsing
- `reassembler.spec.js` — paragraph alignment, line ending normalization, content validation
- `response-parser.spec.js` — parseDelimitedResponse, fallback classification

## Acknowledgments

Word AI Redliner's formatting-preserving reassembly draws on insights from several
excellent open-source projects in the document editing space. We are grateful to
their authors for sharing their work:

- **[docx-redline-js](https://github.com/AnsonLai/docx-redline-js)** by Anson Lai --
  A JavaScript OOXML-level redlining engine whose surgical mode, paragraph
  property cloning, and reconstruction writer patterns informed our approach to
  preserving paragraph and run formatting during document reassembly.

- **[safe-docx](https://github.com/usejunior/safe-docx)** by UseJunior --
  A safe OOXML manipulation library whose paragraph shell cloning, template run
  selection, multi-stage text matching, and run splitting patterns guided our
  thinking on formatting fidelity and style resolution.

- **[adeu](https://pypi.org/project/adeu/)** by Dealfluence Oy (Mikko Korpela, Uzair Ahmed) --
  A Python OOXML redlining tool whose virtual text contract, run coalescing,
  and deep-copy `w:pPr`/`w:rPr` preservation patterns shaped our understanding
  of how to maintain formatting coherence through tracked change operations.

- **[@xmldom/xmldom](https://github.com/xmldom/xmldom)** --
  A W3C-compliant XML DOM implementation whose namespace-aware manipulation
  patterns and whitespace preservation mechanisms underpin correct OOXML handling.

## Security Notes

- **API key storage**: the optional API key is stored in the add-in's
  `localStorage` (plaintext, scoped to the add-in's HTTPS origin). Treat it
  like a browser-saved password: use a key restricted to your LLM endpoint,
  and prefer keyless local backends (Ollama) where possible.
- **LLM output sanitization**: markdown from the LLM is sanitized with
  DOMPurify before being inserted into Word documents, so prompt-injected
  HTML cannot become live markup in a generated summary.
- **HTTPS is mandatory** for Word add-in hosting. The production container
  warns and refuses nothing, but Word itself will block plain HTTP.
- **Static server hardening**: path traversal and malformed-URL crashes are
  handled (400 responses), non-GET/HEAD methods are rejected with 405, and
  the server shuts down gracefully on SIGTERM.
- See `SECURITY.md` for how to report vulnerabilities.

## Licensing

This project is dual-licensed:

- **MIT License** applies to the Word add-in codebase.
- **Apache 2.0 License** applies to the `office-word-diff` library (used as a dependency).

See `LICENSE` and `LICENSE-APACHE` for details.
