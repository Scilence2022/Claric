# Claric ? your redlining scribe for Word

<p align="center">
  <img src="assets/icon-128.png" alt="Claric" width="96" height="96" />
</p>

Claric is an AI-powered Microsoft Word add-in with a chat-driven taskpane. Tell
it what to do ? polish a passage, restyle the headings, add a title, insert a
table, draw an illustration, continue writing, delete empty paragraphs, or
summarize the
comments ? and it routes the request to the right pipeline, stages the result
as a reviewable proposal card, and applies it as native tracked changes.
Instructions work in English or Chinese, against any OpenAI-compatible LLM
backend, local or hosted.

## Features

### Chat Interface & Turn Routing

- Chat-first taskpane: message list, ChatGPT-style composer (rounded input card with a "+" button that opens a slash-command menu, model pill and circular send button in the toolbar row), welcome empty state with slash-command chips, ?/? recall of submitted prompts (terminal-style, draft preserved)
- **Auto-apply toggle** (macOS-style switch in the composer toolbar): when on, staged proposal cards apply automatically as tracked changes the moment the turn settles ? proposals are still staged and reviewable, application just stops waiting for the click; the setting persists (`autoApplyChanges`) and the cross-card apply mutex still guards the write
- Slash-command picker ? type `/` to filter; Enter/Tab/click to select; the send button morphs into Cancel while a run is processing (AbortController)
- Built-in skills: `/copy-edit`, `/check-doc`, `/flag-issues`, `/summarize-contract`, `/industry-overview`, `/storylining`; saved prompts register as custom slash commands
- Intent routing for free text (English and Chinese): edit, format, append, table, illustration, and empty-paragraph cleanup each route to their own pipeline; questions always go to Q&A
- Free text with a non-empty selection ? amendment pipeline with the text as the edit instruction; `selection-first` skills run on the selection when one exists, otherwise on the whole document
- Compound instructions ("add a title, then polish the whole text") are decomposed by an LLM task planner into up to 6 ordered tasks (`insert`, `format`, `edit`, `append`, `table`, `illustration`, `qa`), one proposal card per pipeline
- Ambiguous instructions with no intent keyword are classified by the planner instead of defaulting to Q&A; planning failure falls back gracefully

### Staged Proposals & Per-Change Review

Every document mutation is staged as a proposal card ? nothing is written until you apply it:

- Selection edits show before/after character counts; document-scope runs show per-section citation pills that jump to that section in the document
- One checkbox per change, an inline deleted/inserted diff preview (diff-match-patch semantic cleanup), and a locate button that selects the source text in the document
- Apply writes only the checked changes; Reject discards and cleans up bookmarks
- Honest feedback: skipped sections are reported, no-change echoes are never proposed, and a card that lands nothing settles into a warning state
- Staged ranges are re-anchored at apply time, so applying another card first (e.g. an inserted title) never causes deletions on drifted paragraphs

### AI Redlining

- Word-level tracked changes via the vendored word-diff layer (`src/lib/word-diff/`, based on Apache-2.0 [office-word-diff](https://github.com/yuch85/office-word-diff)) plus a project-original CJK character-level strategy
- Cascading strategies: char diff (CJK) ? token map (preserves run-level formatting: bold, italic, font, color) ? sentence diff ? block replace
- Hardened diff layer: caller-owned change-tracking mode with finally-restore, occurrence-ordered sentence diffs, and Nth-occurrence token resolution

### Whole-Document Processing

- Full-document amendment and commenting: parse ? chunk ? parallel LLM dispatch ? reassemble
- Paragraph-level parsing with style, heading, list, and table detection; token-aware chunking (default 12K tokens per chunk) with heading-based boundaries (H1/H2 trigger splits, H3+ stay coherent), table atomicity, and overlap context
- Tiny trailing chunks are merged into the previous chunk to prevent orphans
- Worker-pool dispatch with auto-tuned concurrency (4 workers for large chunks, 6 for smaller), AbortController cancellation, per-chunk progress with ETA estimation, and retry of failed chunks without re-processing successful ones
- Context extraction: definitions (means / shall mean / is defined as / the-X / hereinafter-X patterns), abbreviations via word-initial matching, and a document outline from the heading hierarchy ? injected as a per-chunk filtered prefix into each LLM system message
- Formatting-preserving reassembly: LCS + word-level similarity paragraph alignment; paragraph properties (styles, numbering, indentation) preserved through paragraph-scoped operations; reverse-order application; bookmark-persisted ranges; line-ending normalization; truncation guard against severely shortened LLM output
- Merged amendment + comment mode: one LLM call returns delimited `===AMENDMENT===` / `===COMMENT===` sections; comments are inserted on bookmarked ranges after all amendments; undelimited responses are treated as amendment-only; comment instructions persist with the prompt data
- Output quality: no-commentary/no-markdown rules appended to prompts, `stripMarkdown()` as a safety net, and `stripThinkTags()` removing reasoning blocks

### Formatting & Insert Ops

- Natural-language formatting instructions ("make this Heading 2", "center all headings") are planned into a strict JSON op allowlist: font (bold, italic, underline, highlight, color, size, ?) and paragraph (style, alignment, spacing, indentation, ?) properties, targeted by substring match, paragraph style, or whole scope
- Insert ops add short structural elements ? "add an article title" composes a title, inserts it at the document start, and styles it with the built-in Title style
- Applied as tracked Formatted revisions; existing text is never rewritten

### Illustration Pipeline

- "Design and insert an illustration" asks the LLM for a self-contained SVG, sanitizes it with DOMPurify (no scripts or `foreignObject`), rasterizes it to PNG via an offscreen canvas, and inserts it as a centered inline picture (? 450 pt wide)
- Position heuristic: header/title-image requests go to the document start, otherwise the end; the proposal card shows an image preview

### Document Append

- "Continue writing" drafts new content against the full document context and stages an append-to-end proposal; Apply inserts it as tracked changes

### Table Creation

- "Insert a 3?3 table at the end of the document" routes to a dedicated table pipeline; explicit dimensions without a content request resolve to an empty grid deterministically ? no LLM call, nothing to hallucinate
- Content-bearing requests ("????????" / "fill it with project data") go through a strict JSON table protocol (`src/lib/table-ops.js`): a rectangular plain-text cell matrix, row/column/cell/character limits, and a dimension check against the instruction (a mismatched model grid rejects the whole proposal)
- The proposal card shows a read-only grid preview; Apply inserts one native Word table at the document start/end or before/after the selection, with the grid style, an optional header row, and AutoFit
- The insertion is recorded as a tracked revision on Word desktop; on hosts without structural-revision support (Word for the web, mobile) it lands untracked with an explicit warning instead of a half-tracked state
- Editing an EXISTING table is a separate route: select its cells and the multi-cell patch protocol (`src/lib/table-patch.js`) stages per-cell edits and row insert/delete ops
- Tables and images enter the conversation as **controllable objects with a tool list**, not raw content: a multi-cell/whole-table selection or an image selection routes to a tool-loop session (`src/lib/table-model.js` / `src/lib/image-model.js`). The table side exposes `get_state` / `set_cell` / `insert_row` / `delete_row` / `merge_cells`; the image side exposes `list_images` / `read_image` (multimodal content reading via the loop) / `design_illustration` / `replace_illustration` / `delete_image` / `resize_image` / `set_alt_text`. Ops are staged as a diffable transaction and applied only on Apply.

### Empty-Paragraph Cleanup

- "Delete empty paragraphs" runs a deterministic Word.js scan (no LLM ? blank paragraphs are invisible to the text pipelines), excluding the final paragraph, table-cell paragraphs, and paragraphs holding inline pictures
- Staged as a proposal card; Apply deletes them as tracked changes and reports the actual count

### Document Summary & Comments

- The `/summarize-contract` skill extracts all document comments, document text, and tracked changes, then generates a formatted Word document in a new file
- `{comments}` inserts structured comment data (author, annotated text, comment text); `{whole document}` extracts the full text with configurable richness ? **Plain** (raw paragraph text), **Headings** (markdown heading markers), **Structured** (headings + list numbering and indentation); `{tracked changes}` extracts revision marks via OOXML parsing
- Generated documents include an annex with numbered source comments; LLM markdown is converted to HTML via [marked](https://github.com/markedjs/marked) and tables render with visible borders
- OOXML tracked-changes parsing uses the browser DOMParser (no external dependencies): handles the `pkg:package` wrapper and `w:proofErr` normalization, pairs adjacent `w:del` + `w:ins` from the same author as replacements, detects move operations (`w:moveFrom` / `w:moveTo`), skips table-row revision markers, queries namespace-aware with prefix fallback, and includes author identity in the LLM-formatted output
- Async comment queue: bookmark-based range persistence, a pending counter with retry-on-error, and WordApi 1.4 detection with graceful degradation

### MCP Tools (Model Context Protocol)

- Connect HTTP (Streamable) MCP tool servers in Settings ? MCP Servers (name, URL, optional bearer token, enable/disable, one-click test)
- `/mcp <instruction>` runs a ReAct tool loop (the same backend-agnostic loop as table/image sessions) over all enabled servers' tools ? tool names are namespaced, results stream into the work log, and the final summary answers in chat
- Read-only contract: MCP tools act on their own external systems and never write to the Word document directly; oversized results are truncated, image results ride the multimodal attachments channel
- Resource tools (`mcp_list_resources` / `mcp_read_resource`) let the model pull reference material from servers mid-loop; "Import prompts" converts an MCP server's prompt templates into slash-command skills; the per-turn tool-loop step budget is configurable (Settings ? MCP Servers)
- CORS-bound servers can be proxied same-origin via the generic `CUSTOM_PROXY_PATH`/`CUSTOM_PROXY_TARGET` pair

### Skill Packages (SKILL.md)

- Import UI lives in Settings ? Skills (its own tab); imported packages appear as slash commands

- Import Claude-style skill packages: YAML frontmatter (name, description, optional `category`/`scope`) over a markdown instruction body ? pasted or loaded from a `.md` file in Settings ? Prompts
- Imported skills appear as slash commands alongside built-ins and saved prompts; they persist in localStorage (`wordAI.skills.imported`, capped at 24) and can be removed in Settings

### Prompt System

- Four independent prompt categories ? context, amendment, comment, summary ? each with full CRUD, per-category activation, and `{selection}` placeholder replacement
- Prompts persist in localStorage across sessions and double as custom slash commands in the chat UI

### Model Activity, Work Log & Streaming

- Each turn shows a collapsible work log ("Worked for Ns ? M steps") and a live model-activity region streaming reasoning (dimmed) and output tokens per section
- Model activity auto-scrolls while pinned to the bottom; scrolling up disengages, scrolling back re-engages
- Chat answers and pipelines stream token-by-token via OpenAI-compatible SSE; an idle timeout resets on every chunk, so long generations survive and only stalled streams abort; automatic fallback to non-streaming
- A live selection preview chip sits above the input bar; the current selection is injected into Q&A prompts as focused context

### LLM Backends

- Providers: Ollama, vLLM, DeepSeek, Zhipu GLM, Moonshot Kimi, MiniMax (international + China sites), and Custom (any OpenAI-compatible endpoint)
- Unified OpenAI-compatible chat API; per-provider API prefix handled automatically (GLM uses `/api/paas/v4`)
- Cloud provider defaults are origin-adaptive: on statically hosted installs (marketplace/Pages) they point at the absolute API origins (e.g. `https://api.deepseek.com`) and are called directly ? all five providers send CORS headers for public origins; behind the local dev/production server they point at same-origin proxy paths, because those providers refuse CORS for localhost/private-IP origins
- Local models (Ollama/vLLM) default to same-origin proxy paths (`/ollama`, `/vllm`) served by the dev/production server ? required because an HTTPS page cannot call `http://localhost` (mixed-content blocking) and to avoid per-user CORS setup
- Typeable model field with a refreshable suggestion list (Refresh re-queries the provider's models endpoint); configurable endpoint URL and optional API key per provider; Track Changes and Line Diff toggles

### Settings & UX

- All provider, extraction, and prompt settings live in a slide-over panel with auto-save (plus an explicit Save button for visible confirmation)
- Model pill under the input bar shows the active provider:model and opens settings on click
- Activity log in a slim drawer; connection status and comment-pending indicators in the status bar

## Setup

There are **two ways** to run this add-in:

| Method | Best for | Requirements |
|--------|----------|--------------|
| **Docker** | Quick setup, no Node.js needed | Docker, Docker Compose |
| **npm** | Development, customization | Node.js 22+ |

Both methods require HTTPS certificates trusted by the machine running Word.

---

## Deployment Routes: Static vs Local Server

Beyond *how you host the server*, the add-in UI itself runs from one of two
origins. Both are first-class; switching is one command each.

| | **Static route** (GitHub Pages / Microsoft Marketplace) | **Local-server route** (Docker / npm dev) |
|---|---|---|
| Taskpane served from | `https://scilence2022.github.io/claric-addin/` | `https://localhost:<HOST_PORT>` (or a LAN IP) |
| Cloud providers (DeepSeek / GLM / Kimi / MiniMax) | Direct CORS calls to the provider's API origin ? no server needed (API key required) | Same-origin proxy paths (`/deepseek`, ?) ? the providers refuse CORS for localhost/private-IP origins (verified), so direct calls cannot work here |
| Local models (Ollama / vLLM) | Not reachable: an HTTPS page cannot call `http://localhost` (mixed-content blocking; WebKit has no exemption ? bugs.webkit.org 171934/173161). Options: an HTTPS relay in front of Ollama (`OLLAMA_ORIGINS` for CORS), or use the local-server route | ? Work out of the box via the default proxy paths |
| Backend server required | None | `docker compose up -d` or `npm start` |
| Typical use | Marketplace submission, everyday cloud-LLM use | Local development (hot reload), local-AI setups |

Provider defaults are **origin-adaptive**: on static hosts cloud presets
default to their absolute API origins; behind the local server they default
to the proxy paths. Both routes are zero-config for their supported
backends.

### Switching routes

| Command | Effect |
|---------|--------|
| `npm run manifest:local` | Point `manifest.xml` at the local server (`HOST=localhost`, `PORT=3001`) and regenerate it |
| `npm run manifest:store` | Point `manifest.xml` at the GitHub Pages host (`HOST=scilence2022.github.io`, base path `/claric-addin`) and regenerate it |
| `npm run sideload` | Sideload the current `manifest.xml`: into Word for Mac's `wef/` folder, or via the Windows installer (developer-add-in registry + launch document) |
| `npm run sideload:remove` | Remove the sideloaded manifest from Word for Mac |
| `npm run publish:addin` | Build the production bundle, switch to the store manifest, push `dist/` + `manifest.xml` to the Pages repo (`Scilence2022/claric-addin`); GitHub Pages rebuilds in ~30 s |

A typical local session: `npm start` (or `docker compose up -d`) ?
`npm run manifest:local` ? `npm run sideload` ? restart Word.
A typical release: bump `package.json` ? `npm run publish:addin` ?
`npm run sideload` ? restart Word.

### Route-specific notes

- **Settings are per-origin.** Provider choices, endpoints, and API keys live
  in the add-in's localStorage under the serving origin
  (`https://localhost:3001` vs `https://scilence2022.github.io`). Switching
  routes starts from a fresh Settings state ? re-enter the API key once per
  route.
- **Static installs and local models.** `http://localhost` endpoints are
  blocked by Word's WebView from an HTTPS page. The supported alternatives
  are an HTTPS reverse proxy in front of Ollama (plus `OLLAMA_ORIGINS`
  including the add-in's origin) or a reachable relay of the production
  container ? enter its absolute URL as the endpoint.
- **Why cloud defaults differ per route (verified, not configurable).** The
  five providers reflect `Access-Control-Allow-Origin` for public origins
  (github.io, arbitrary domains) but emit **no CORS headers for localhost
  or private-IP origins** ? a common provider-side policy against local
  network attacks. Hence: static installs call them directly; the local
  server must relay. The same asymmetry plus mixed-content blocking rules
  out direct `http://localhost` calls for local models.
- **Manifest `BASE_PATH`.** Defaults to `/claric-addin` when
  `HOST=scilence2022.github.io` (the Pages repository path); empty for any
  other host. Override explicitly in `.env` if you publish under a
  different repo name.

---

## Option A: Docker (Recommended for Quick Setup)

### Prerequisites

- Docker and Docker Compose
- HTTPS certificate files (see [Create HTTPS Certificates](#create-https-certificates))

### Step-by-Step

1. **Clone the repository**

```bash
git clone https://github.com/Scilence2022/Claric.git
cd Claric
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
git clone https://github.com/Scilence2022/Claric.git
cd Claric
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
4. Navigate to **Trusted Root Certification Authorities** ? **Certificates**.
5. Right-click ? **All Tasks** ? **Import...**
6. Select the `.crt` file and finish the wizard.

**If you used mkcert**, you can install the mkcert root CA on Windows instead:

- Copy the root CA from the server machine (find it via `mkcert -CAROOT`)
- Import it into **Trusted Root Certification Authorities**

---

## Sideload the Add-in

### Word on Windows

**Method 1: One-click installer (recommended)**

No admin rights, no Node.js. Uses the per-user developer-add-in registry
(`HKCU\...\16.0\Wef\Developer`) ? the same mechanism `office-addin-dev-settings`
employs ? so it also works on consumer Microsoft 365 builds where the
"Upload My Add-in" entry no longer appears.

- **Static build (GitHub Pages)** ? run in PowerShell:

  ```powershell
  irm https://raw.githubusercontent.com/Scilence2022/Claric/main/installer/windows/Install-Claric.ps1 | iex
  ```

- **Self-hosted build (Docker / npm)** ? from a checkout, after generating the
  manifest for your server:

  ```
  npm run install:windows        # wraps installer/windows/Install-Claric.ps1 -ManifestPath manifest.xml
  ```

The installer copies the manifest to `%LOCALAPPDATA%\ClaricAddin\`, registers
it, builds `Claric-Launch.docx`, and opens Word with the taskpane mounted.
Reopen that document whenever you need Claric, or find it under
**Insert ? Get Add-ins ? My Add-ins**. Undo with `Uninstall-Claric.ps1` or
`npm run uninstall:windows`. Details and troubleshooting:
[installer/windows/README.md](installer/windows/README.md).

**Method 2: Add from file** (if your Word build still offers it)

1. Open Word ? **Insert** ? **Get Add-ins** ? **My Add-ins**.
2. Click **Add a custom add-in** ? **Add from file...**.
3. Select `manifest.xml` and confirm.

**Method 3: Network shared folder (Windows only)**

1. Create a shared folder and note the network path (a real UNC share ?
   a plain local path such as `C:\...` is not picked up).
2. In Word: **File** ? **Options** ? **Trust Center** ? **Trust Center Settings** ?
   **Trusted Add-in Catalogs** ? **Add catalog** (check **Show in Menu**).
3. Copy `manifest.xml` into the shared folder.
4. In Word: **Home** ? **Add-ins** ? **Advanced** ? **Shared Folder** ? select the add-in ? **Add**.

For full details, see the [Microsoft sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

### Word on Mac

Mac's Word has no "Add from file" dialog -- sideloading uses the `wef`
container folder (see [Microsoft's Mac sideloading guide](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)):

1. Make sure the server is running (`docker compose up -d`).
2. Open Finder and press ?+Shift+G, then enter:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   (create the `wef` folder if it does not exist).
3. Copy the `manifest.xml` into that folder. The file lives on the server:
   `https://<HOST>:<HOST_PORT>/manifest.xml` (save it from a browser).
4. Fully quit Word (?Q) and reopen it -- loaded only at launch.
5. Go to the **Home** tab ? **Add-ins** ? select **Claric**.
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
| New `*_PROXY_PATH` in `.env` does not take effect after `docker compose restart` | `restart` keeps the original env-file resolution. Run `docker compose down && docker compose up -d` to re-parse `.env` (no image rebuild needed) |
| LLM proxy routes log only `/ollama` and `/vllm` (or none) when you expected more | Provider proxy routes are **disabled by default** in production; only paths explicitly set in `.env` are wired. Add `DEEPSEEK_PROXY_PATH=/deepseek` etc. and recreate the container (`down` + `up -d`) |
| `curl https://localhost:<PORT>` fails with SSL "wrong version number" or no response | Another local server is bound to that port (e.g. a Next.js dev server on `:3000`). Pick a free port via `HOST_PORT` in `.env` (e.g. `3001`) and re-run `up -d`. The container still listens internally on `3000` |
| Browser/Word refuses the cert with a hostname mismatch | The `server.pem` must cover the hostname Word connects with. Regenerate with `mkcert <your-host-or-ip>` and copy to `server.pem`/`server-key.pem` (SAN must include the host) |
| `docker compose build` fails on `npm ci` in the builder stage | Usually a transient npm registry hiccup while pulling 800+ packages. Re-run `docker compose build --no-cache` once; if it persists, check `package-lock.json` vs `package.json` consistency |
| Model connection fails on the static route (Pages/marketplace) with Ollama/vLLM | A static HTTPS page cannot reach `http://localhost` (mixed-content blocking). Either switch to the local-server route (`npm run manifest:local`), or front the backend with HTTPS + CORS (`OLLAMA_ORIGINS`) and enter its absolute URL as the endpoint |
| Add-in still shows old behavior after a Pages publish | GitHub Pages assets are cached ~10 minutes. Fully quit and reopen Word; if still stale, delete `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/WebKit` while Word is closed |

### Deployment footnotes

- **Editing `.env` requires a full recreate, not a restart.** Docker Compose resolves `env_file` at container creation and caches the result; `restart` reuses the cached config. Always use `docker compose down && docker compose up -d` after editing `.env`. The image is not rebuilt, only the container is recreated.
- **`HOST_PORT` vs `PORT`.** The container always binds `PORT=3000` internally. `HOST_PORT` is the host-side port published by compose. Pick `HOST_PORT` carefully to avoid collisions with other local services (Next.js, dev servers, etc.).
- **LLM proxy default behavior changed in 0.5.0.** All `*_PROXY_PATH` env vars now default to empty (= disabled). To re-enable a provider, set both `*_PROXY_PATH` and (if non-default) `*_PROXY_TARGET` in `.env`.
- **API keys live client-side, not in `.env`.** Cloud provider keys (DeepSeek / GLM / Kimi / MiniMax) are entered in the add-in's Settings UI and stored in localStorage. The server-side proxy only forwards requests; it does not read or inject keys.
- **Rebuild strategy.** If the diff is `src/**` + tests + docs only ? `docker compose build` (incremental, ~seconds). If `package.json`, `package-lock.json`, `webpack.config.cjs`, or `Dockerfile` changed ? `docker compose build --no-cache` (~minutes).

---

## Environment Variables

### All deployments

> LLM proxy paths (`*_PROXY_PATH`) are **disabled by default in production**
> (`scripts/docker-server.cjs`): set a path explicitly (e.g. `/ollama`) to
> enable that provider's route. The defaults below apply to the webpack dev
> server, where they are enabled for local development.

| Variable | Default | Description |
|----------|---------|-------------|
| `HOST` | `localhost` | Hostname for manifest URLs (must be reachable from Word) |
| `PORT` | `3000` | Port for manifest URLs / server listen port |
| `PROTOCOL` | `https` | Protocol for manifest URLs and the server |
| `SSL_CERT_FILE` | `server.pem` | Path to SSL certificate |
| `SSL_KEY_FILE` | `server-key.pem` | Path to SSL private key |
| `ADDIN_GUID` | *(generated)* | Stable add-in identity; auto-generated and persisted to `.manifest-guid` on first manifest generation. Pin it in Docker so container recreation keeps the identity. |
| `DISPLAY_NAME` | `Claric ? AI Writing & Editing Assistant for Word` | Display name shown in Word and the store (store validation policy: no Microsoft/Office/Word branding as the product name; "X for Word" suffixes are fine) |
| `SUPPORT_URL` | *(add-in root URL)* | Support page required by store validation ? point it at a real support page, e.g. GitHub issues |
| `APP_DOMAINS` | *(none)* | Comma-separated extra domains for the manifest's `AppDomains` element (the SourceLocation domain is implicitly trusted) |
| `OLLAMA_PROXY_PATH` | *(disabled)* / `/ollama` (dev) | Proxy path for the Ollama backend (empty disables) |
| `OLLAMA_PROXY_TARGET` | `http://localhost:11434` | Upstream Ollama base URL (`http://host.docker.internal:11434` in Docker) |
| `VLLM_PROXY_PATH` | *(disabled)* / `/vllm` (dev) | Proxy path for the vLLM backend (empty disables) |
| `VLLM_PROXY_TARGET` | `http://localhost:8026` | Upstream vLLM base URL (`http://host.docker.internal:8026` in Docker) |
| `DEEPSEEK_PROXY_PATH` | *(disabled)* / `/deepseek` (dev) | Proxy path for DeepSeek (empty disables) |
| `DEEPSEEK_PROXY_TARGET` | `https://api.deepseek.com` | Upstream DeepSeek API origin |
| `GLM_PROXY_PATH` | *(disabled)* / `/glm` (dev) | Proxy path for Zhipu GLM (empty disables) |
| `GLM_PROXY_TARGET` | `https://open.bigmodel.cn` | Upstream Zhipu GLM API origin |
| `KIMI_PROXY_PATH` | *(disabled)* / `/kimi` (dev) | Proxy path for Moonshot Kimi (empty disables) |
| `KIMI_PROXY_TARGET` | `https://api.moonshot.cn` | Upstream Moonshot Kimi API origin |
| `MINIMAX_PROXY_PATH` | *(disabled)* / `/minimax` (dev) | Proxy path for MiniMax international (empty disables) |
| `MINIMAX_PROXY_TARGET` | `https://api.minimax.io` | Upstream MiniMax API origin |
| `MINIMAX_CN_PROXY_PATH` | *(disabled)* / `/minimax-cn` (dev) | Proxy path for MiniMax China (empty disables) |
| `MINIMAX_CN_PROXY_TARGET` | `https://api.minimaxi.com` | Upstream MiniMax China API origin |
| `LLM_PROXY_TIMEOUT_MS` | `300000` | Proxy upstream timeout in ms |

### Dev server only (webpack)

| Variable | Default | Description |
|----------|---------|-------------|
| `DEV_SERVER_HOST` | `127.0.0.1` | Host to bind webpack dev server (`0.0.0.0` exposes it to the network) |
| `DEV_SERVER_PORT` | `3000` | Port for webpack dev server |
| `ENABLE_DEV_ENDPOINTS` | *(off)* | Set `true` to register the dev-only E2E/coding-agent endpoints (`/log`, `/api/e2e-loop/*`, `/api/test-cases`, `/api/prompts`) ? they write files and use wildcard CORS; see `scripts/dev-e2e-middlewares.cjs` |
| `LLM_PROXY_TLS_VERIFY` | `true` | Set `false` only for a local LLM backend with a self-signed certificate |
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

Build the production image locally:

```bash
docker build -t claric .
```

The image is a three-stage build on `node:22-alpine`: runtime layers contain
production dependencies only, run as the non-root `node` user, and expose a
`/healthz` endpoint used by the Docker `HEALTHCHECK`.

---

## Project Structure

See `ARCHITECTURE.md` for details.

## Testing

```bash
npm test          # ~1300 tests across 60 suites
npm run lint      # ESLint (flat config)
npm run build     # webpack production build
npm run verify    # lint + test + build (same as CI)
```

Test suites cover:
- `prompt-state.spec.js` ? PromptManager CRUD, activation, persistence, summary category
- `prompt-persistence.spec.js` ? localStorage round-trip, migration, edge cases
- `prompt-composition.spec.js` ? composeMessages, composeSummaryMessages, placeholder replacement, output rules
- `comment-extractor.spec.js` ? extractAllComments, extractDocumentStructured, estimateTokenCount, extractTrackedChanges (OOXML parsing)
- `comments-on-range.spec.js` ? extractCommentsOnRange
- `document-generator.spec.js` ? buildSummaryHtml (markdown conversion, table borders, escaping), createSummaryDocument (Word API)
- `comment-queue.spec.js` ? CommentQueue state management, bookmark naming
- `llm-client.spec.js` ? sendPrompt, sendMessages, stripThinkTags, stripMarkdown, testConnection
- `llm-stream.spec.js` ? sendMessagesStream SSE parsing, reasoning demux, [DONE] terminator, non-SSE fallback, abort, idle timeout
- `document-parser.spec.js` ? parseDocument, paragraph extraction, heading detection, style mapping
- `document-chunker.spec.js` ? chunkDocument, heading-based splitting, overlap, token limits
- `context-extractor.spec.js` ? extractContext, definitions, abbreviations, outline generation
- `orchestrator.spec.js` ? processChunksParallel, concurrency, cancellation, merged mode parsing, streaming tokens
- `reassembler.spec.js` ? paragraph alignment, line ending normalization, content validation, staged-range re-anchoring, blank-paragraph handling
- `response-parser.spec.js` ? parseDelimitedResponse, fallback classification
- `format-ops.spec.js` ? parseFormatOps allowlist sanitizing, insert ops, describeFormatOp
- `table-ops.spec.js` ? table creation spec inference (EN/ZH dimensions), creation prompt, strict response parsing/validation, limits
- `table-patch.spec.js` ? multi-cell patch prompt, JSON parsing/validation, bounds/row-op rules
- `word-actions-table.spec.js` / `word-actions-table-create.spec.js` ? table selection patch route (ordering, tracking split, staleness guards) and table creation route (empty-grid fast path, constrained generation, insertion positions, platform tracking split)
- `illustration.spec.js` ? parseIllustration, SVG sanitizing, dimensions, position heuristic
- `task-planner.spec.js` ? parsePlan caps/truncation, buildPlanPrompt
- `providers.spec.js` ? provider preset catalog, MiniMax pair, default config
- `char-diff.spec.js` ? CJK detection, computeCharEdits, applyCharDiffStrategy
- `word-diff.spec.js` ? word/sentence diff modes, sliceSearchPieces
- `skills.spec.js` ? built-in skill registry, resolveSkill parsing, custom skill registration
- `conversation.spec.js` ? chat turn routing (skill / selection edit / doc edit / format / append / table / illustration / compound / Q&A), staged proposals, selective apply, honest warnings, concurrency guard, cancel
- `proposal-card.spec.js` ? per-change checkboxes, inline diff, locate, selective apply, warning/error states
- `chat-view.spec.js` ? model activity auto-scroll follow/disengage behavior
- `config-persistence.spec.js` ? normalizeConfig validation and legacy migration
- `selection-with-comments.spec.js` ? comment anchor splicing into selection OOXML
- `generate-manifest.spec.js` ? manifest generation from template + .env

## Acknowledgments

Claric is a fork of **[word-ai-redliner](https://github.com/yuch85/word-ai-redliner)**
by yuch85 ? its original architecture (the IR pipeline, chunking, and
formatting-preserving reassembly) is the foundation this codebase builds on.

Beyond that, Claric's reassembly also draws on insights from several excellent
open-source projects in the document editing space. We are grateful to their
authors for sharing their work:

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
- **Apache 2.0 License** applies to the diff strategies vendored from `office-word-diff` in `src/lib/word-diff/` (see `LICENSE`/`NOTICE` there).

See `LICENSE` and `LICENSE-APACHE` for details.
