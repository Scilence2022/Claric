# Claric — Codebase Analysis Report & Improvement Plan

**Date:** 2026-08-25
**Scope:** Full-repository review of the Claric Microsoft Word add-in
**Status:** Analysis complete (static review). Verification commands (lint/test/build/audit) could not be executed in this environment — see [§2 Methodology](#2-methodology-and-caveats).

---

## 1. Executive Summary

Claric is, on the whole, a **well-architected and unusually well-documented** codebase. It has a clear separation between pure logic (`src/lib/`), orchestration (`src/taskpane/`), and presentation (`src/taskpane/ui/`); a rigorous, heavily-documented turn-routing design; a strong unit-test suite (29 suites, ~700 tests); a hardended production server; and a genuine security consciousness (DOMPurify sanitization of LLM output, XML escaping, a SECURITY.md that is honest about trade-offs).

The most significant problems are **not** in the core application logic, which is solid, but in three surrounding areas:

1. **Development/test infrastructure is entangled with the production build config.** The `webpack.config.cjs` file is 752 lines long, of which roughly 520 lines (~70%) implement E2E-test and coding-agent helper HTTP endpoints (`/log`, `/logs`, `/api/trace-log`, `/api/fix-log`, `/api/e2e-loop/*`, `/api/test-cases`, `/api/prompts`) served with wildcard CORS. This is a separation-of-concerns and attack-surface problem.
2. **Dead code and an incomplete project rename** ("Word AI Redliner" → "Claric") create maintenance drag and brand inconsistency, including a stale CI/Docker image tag.
3. **An abandoned runtime dependency** (`diff-match-patch@1.0.5`, unmaintained since ~2017) sits in the hot path of the core red-lining feature.

None of these are drop-everything emergencies, but they are the highest-leverage improvements available.

---

## 2. Methodology and Caveats

- This review is **static**: it was performed by reading source, configuration, and test files and by cross-referencing `README.md`, `ARCHITECTURE.md`, and `SECURITY.md`.
- **Caveat — environment limits:** the review environment could not fork external processes (every `git`/`npm`/`ls` invocation returned `fork: Resource temporarily unavailable`). Consequently:
  - **A new `git worktree` could not be created**, so the analysis was performed against the working tree (read-only), and this report is written to `docs/` in the main tree.
  - **`npm run lint`, `npm test`, `npm run build`, `npm audit`, and `npm outdated` could not be executed.** Findings that depend on tooling (e.g., dependency-CVE status) are marked as "to verify" rather than asserted.
- Dependency versions were verified by reading `package-lock.json` directly.
- To create the worktree yourself (once the environment recovers), run:
  ```
  git worktree add ../Claric-analysis main
  ```

Severity key: **Critical** (exploitable/blocking) · **High** (should fix soon) · **Medium** (fix this cycle) · **Low** (polish).

---

## 3. Codebase Overview

| Attribute | Value |
|---|---|
| Product | AI-powered Word add-in (redlining, summarization, comment analysis) |
| Runtime | Office.js taskpane (WebView2), ES modules, plain JS + JSDoc (no TypeScript) |
| Node | >= 22 (`.nvmrc` = `22`) |
| Build | Webpack 5, `npm run build` |
| Tests | Jest (babel-jest), jsdom + node environments |
| Lint | ESLint 9 flat config |
| Runtime deps | `diff-match-patch@1.0.5`, `dompurify@3.4.14`, `dotenv@17.2.3`, `marked@17.0.4` |
| Serving | webpack-dev-server (dev) · `scripts/docker-server.cjs` (prod, non-root node:22-alpine) |
| CI | GitHub Actions (`ci.yml`): lint + test + build + docker build |

Key modules: `src/lib/` (pure logic — llm-client, orchestrator, reassembler, chunker, parser, format-ops, illustration, word-diff/*), `src/taskpane/` (conversation routing, word-actions pipelines, ui/*), `scripts/` (manifest generation, production server).

---

## 4. Findings by Dimension

### 4.1 Architecture & Maintainability

| # | Sev | Finding | Evidence | Recommendation |
|---|-----|---------|----------|----------------|
| A1 | **High** | Test/coding-agent endpoints are baked into the production-relevant `webpack.config.cjs`. ~520 of 752 lines are dev-only E2E logging, loop-control, test-case CRUD, and prompt CRUD routes, all with `Access-Control-Allow-Origin: *`. | `webpack.config.cjs:213-743` (`setupMiddlewares`) | Extract into a separate dev middleware module/plugin (or a `webpack.dev.cjs`) that is never part of the add-in build; gate behind an `ENABLE_DEV_ENDPOINTS` flag; default off. |
| A2 | **Medium** | Dead / legacy modules retained only for their tests. | `src/lib/panel-actions.js` (header: "Legacy frozen enums (kept for tests)"), `src/lib/structure-model.js` (`ParagraphBlock` "Token Map" model superseded by `word-diff/`), exercise `tests/panel-actions.spec.js`, `tests/structure-model` references. | Delete both modules and their specs; if the enum constants are still useful elsewhere, move them next to the consumer or document the single source of truth. |
| A3 | **Medium** | Incomplete "Word AI Redliner" → "Claric" rename leaves stale branding and a stale image/tag. | `docker-compose.yml:13` (`ghcr.io/yuch85/word-ai-redliner:0.4.0` + "rename the GHCR repo … update this tag" TODO), `.github/workflows/ci.yml:44` (`docker build -t word-ai-redliner:ci`), `.env.example:2` ("Word AI Redliner - Environment Configuration"), `docs/word-ai-redliner.gif`. Version drift: docker image tag is 0.4.0 vs `package.json` 0.5.0. | Finish the rename; bump the image tag to match `package.json`; rename the gif. |
| A4 | **Low** | Logic duplication between `verify-word-api.js` and the word-diff layer. | `src/scripts/verify-word-api.js:333` duplicates `const token = tokensAfterDeletes[currentTokenIdx]` from `src/lib/word-diff/token-map.js:149`. | Move shared token-map helpers into `word-diff/` and have the script import them, or delete the script if it is a one-off diagnostic. |
| A5 | **Low** | Test-only re-export in the production bootstrap. | `src/taskpane/taskpane.js:37` `export { normalizeConfig } from './app-state.js'` (documented "test seam"). | Point the test at `app-state.js` directly and drop the re-export. |
| A6 | **Info** | Intent routing relies on a single large, hand-maintained regex set (EN + ZH) per intent family. This is well-documented and unit-tested, but is brittle to language drift. | `src/taskpane/conversation.js:53-158`. | Keep; consider a data-driven intent table as a long-term refactor; add a regression corpus of real user phrasings. |

### 4.2 Code Quality

| # | Sev | Finding | Evidence | Recommendation |
|---|-----|---------|----------|----------------|
| Q1 | **Low** | The lazy DOMPurify factory (`getPurifier`) is duplicated verbatim in two modules. | `src/lib/document-generator.js:19-26`, `src/lib/illustration.js:16-23` (identical) | Extract a shared `src/lib/dompurify.js` (or `sanitize.js`) util. |
| Q2 | **Low** | HTML/XML escaping helpers are near-duplicates in different layers. | `escapeHtml` in `document-generator.js:64-71` vs `escapeXml` in `scripts/generate-manifest.cjs:34-41`. | Consolidate into one escaped-entities helper (noting `escapeXml` also escapes `'`). |
| Q3 | **Low** | No static type checking — ~50 modules and ~700 tests rely on JSDoc alone; type errors surface at runtime in a WebView2 environment that is hard to debug. | whole repo (ES modules, `@types/office-js` used only for editor hints). | Incremental `tsc --checkJs` / `@ts-check` adoption, or a full TS migration plan (see §5 P2). |
| Q4 | **Info** | Magic numbers/constants are inconsistently centralized (good: `MAX_SVG_CHARS`, `MAX_INSERT_CHARS`, `FONT_*_KEYS` in `format-ops.js`; scattered: 5-min proxy timeout repeated in `webpack.config.cjs` and `docker-server.cjs`). | `webpack.config.cjs:56` vs `scripts/docker-server.cjs:53`. | Centralize shared limits/timeouts in a small constants module. |

### 4.3 Security

| # | Sev | Finding | Evidence | Recommendation |
|---|-----|---------|----------|----------------|
| S1 | **Medium** | Dev server binds all interfaces with permissive CORS and a dozen write-capable test endpoints — an open attack surface if `npm start` is ever left exposed beyond the dev machine. | `webpack.config.cjs:205-211` (`host: 0.0.0.0`, `allowedHosts: 'all'`, header `*`), plus `POST /api/test-cases`, `POST /api/prompts` (file writes), all `Access-Control-Allow-Origin: *`. SECURITY.md:39-41 already flags this as dev-only. | Default `DEV_SERVER_HOST=127.0.0.1`; move endpoints behind the flag from A1; intentionally documented exceptions only. |
| S2 | **Medium** | The LLM proxy acts as an unauthenticated relay to admin-configured upstreams, forwarding all headers (incl. `Authorization`) with TLS verification disabled in dev. | `webpack.config.cjs:78` (`secure: false`), `scripts/docker-server.cjs:245-291` (forwards headers except hop-by-hop; no auth/rate-limit), `SECURITY.md:30-33`. | On the dev server, remove `secure:false` (or scope it); document the proxy as trusted-intranet-only; consider a per-path allowlist and a minimal no-op auth header strip for cloud paths. |
| S3 | **Low** | API keys stored client-side in `localStorage` and sent to a user-configurable URL. | `src/lib/llm-client.js:244-252`, `SECURITY.md:35-38` (documented trade-off). | Acceptable for this design; add a visible warning when a cloud provider key is set with a non-HTTPS endpoint URL. |
| S4 | **Low** | `process.removeAllListeners('SIGINT'/'SIGTERM')` in the dev config wipes any other installed signal handlers. | `webpack.config.cjs:459-460`. | Replace with named handlers + `process.off` of only the ones added, or fold into the extracted dev module. |
| S5 | **Info** | Prompt injection surface: untrusted document text enters LLM prompts; defense-in-depth is DOMPurify (markdown + SVG) and the strict format-op allowlist. No content-type-policy on `marked`. | `document-generator.js:53-57`, `format-ops.js:110-246`, `illustration.js:112-117`. | Strong baseline; consider a `FORBID_TAGS`/`FORBID_ATTR` and `ALLOWED_URI_REGEXP` tightening in the DOMPurify configs for defense-in-depth. |

### 4.4 Data Handling

| # | Sev | Finding | Evidence | Recommendation |
|---|-----|---------|----------|----------------|
| D1 | **Info** | Full document text, comments, and tracked changes are sent to the configured LLM (self-hosted or cloud). | `SECURITY.md:44-47`, `README.md` summary flow. | Already documented; ensure this is surfaced in the UI near cloud-provider selection (not only in docs). |

### 4.5 Testing

| # | Sev | Finding | Evidence | Recommendation |
|---|-----|---------|----------|----------------|
| T1 | **Medium** | Coverage metric excludes UI and orchestration code, understating the untested surface. | `jest.config.cjs:16-18` (`collectCoverageFrom: ['src/lib/**/*.js']`) — excludes `src/taskpane/**`, `src/commands/**`, and some `word-diff` paths. | Expand `collectCoverageFrom` to the full `src/` and set a baseline; add targeted UI/orchestration tests. |
| T2 | **Low** | Jest environment declared `node`, but many specs and the `jest-environment-jsdom` dependency imply DOM needs. | `jest.config.cjs:2` (`testEnvironment: 'node'`) vs `package.json` (`jest-environment-jsdom`) and `eslint.config.cjs` ("jsdom environment"). | Set an explicit per-suite `@jest-environment` strategy and default to jsdom only where needed; remove the ambiguity. |
| T3 | **Low** | Stale RED-state scaffolding in one spec. | `tests/selection-with-comments.spec.js:3-41` — header claims `src/lib/selection-with-comments.js` "does not exist yet" and wraps `require` in a try/catch `loadError` guard, but the module now exists. | Remove the `loadError` guard and stale header comments; keep the fixtures. |
| T4 | **Info** | No integration/E2E tests run in CI; the E2E harness exists but is untethered to CI. | `webpack.config.cjs` E2E endpoints vs `ci.yml` (unit only). | Either wire a subset of E2E into CI (Windows/Word runner) or delete the harness if unused (ties to A1). |

### 4.6 Build, DevOps & Dependencies

| # | Sev | Finding | Evidence | Recommendation |
|---|-----|---------|----------|----------------|
| B1 | **High** | `diff-match-patch@1.0.5` is abandoned (last published ~2017) and sits in the core diff hot path. | `package-lock.json:5098-5102`. | Replace with a maintained library (`diff` npm package, or `jsdiff`) or vendor and pin it; add a regression test around the semantic-cleanup behavior. |
| B2 | **Medium** | CI performs no dependency-vulnerability scanning. | `ci.yml` (no `npm audit`, no OSV/CodeQL/Dependabot). | Add `npm audit --omit=dev` as a CI gate and enable Dependabot/OSV; **verify current CVE status** (could not be run here). |
| B3 | **Low** | Docker CI job only builds (no push, no scan) and uses the stale image name. | `ci.yml:37-44`. | Extend to tag+push on release, run `docker scan`/`trivy`; fix the name. |
| B4 | **Low** | `marked@17.0.4` and `dompurify@3.4.14` are current (no action). | `package-lock.json:8166-8176, 5177-5185`. | — (recorded to avoid false-positive chasing). |
| B5 | **Low** | `.nvmrc` unpinned major ("22"). | `.nvmrc:1`. | Pin an exact patch (e.g. `22.12.0`) for reproducible builds; keep `engines` range for consumers. |
| B6 | **Info** | Production hardening is strong: `safeJoin` path-traversal guard, hop-by-hop header stripping, graceful shutdown, malformed-percent decode rejection, healthz probe. | `scripts/docker-server.cjs:82-165, 227-291, 352-392`. | Keep; no change needed (positive finding). |

---

## 5. Prioritized Improvement Plan

### P0 — Do first (security/dependency hygiene)

1. **Replace or vendor `diff-match-patch`** (B1). Add a focused regression test for diff semantic cleanup before swapping.
2. **Decouple dev/test endpoints from the add-in build** (A1/S1): extract the `setupMiddlewares` block into a separate, flag-gated dev module; default `DEV_SERVER_HOST=127.0.0.1`; confirm nothing production-facing depends on it.

### P1 — Next cycle

3. **Finish the rename** and fix the stale image/tag (A3): `docker-compose.yml`, `ci.yml`, `.env.example`, `docs/*.gif`.
4. **Delete dead code** and its specs (A2, T3): `panel-actions.js`, `structure-model.js`, stale RED guard in `selection-with-comments.spec.js`.
5. **Add dependency scanning to CI** and run a one-off `npm audit` (B2).

### P2 — Medium-term hardening & tooling

6. **Honest coverage**: broaden Jest coverage scope (T1); align Jest environment config (T2).
7. **Tighten DOMPurify configs** with `FORBID_TAGS`/`FORBID_ATTR`/`ALLOWED_URI_REGEXP` (S5).
8. **Deduplicate shared utilities** (Q1, Q2): `getPurifier`, escape helpers, shared constants (Q4).
9. **Introduce type safety incrementally** (`@ts-check` → `tsc --checkJs` → full TS) (Q3).

### P3 — Polish

10. Replace `process.removeAllListeners` usage (S4); fix `secure:false` scoping (S2).
11. Wire the existing E2E harness into CI or remove it (T4); resolve `verify-word-api.js` duplication (A4); remove the test-seam re-export (A5).
12. Pin `.nvmrc` (B5); extend Docker CI to push + scan (B3).

---

## 6. Strengths (worth preserving)

- **Clean layering** — pure `lib/` modules are DOM- and Word-free, making them trivially unit-testable.
- **Exceptional documentation** — JSDoc on nearly every function; `ARCHITECTURE.md` accurately mirrors the code; `SECURITY.md` is candid about real trade-offs.
- **Defensive-by-design pipelines** — strict JSON allowlists (`format-ops.js`), DOMPurify on all LLM-generated markup/SVG, token-budgeted chunking, crash-safe production server.
- **Strong test suite** with dependency injection seams (`createConversation` deps) enabling ui-free orchestration tests.
- **Honest UX contracts** — staged proposals, "nothing written until applied", no-change warnings — reflected in the code's gating logic.
- **Correct secret hygiene** — `.gitignore` properly excludes `.env`, `*.pem`, `*.key`, `manifest.xml`, and `dist/`; no secrets are tracked.

---

## 7. Outstanding Questions for the Author

1. Is the E2E/coding-agent harness (in `webpack.config.cjs`) still actively used? If not, deleting it resolves A1, S1, S4, and T4 at once.
2. Is `src/scripts/verify-word-api.js` a maintained diagnostic or a one-off? This decides A4 (dedupe vs delete).
3. Is there appetite for a TypeScript migration, or is JSDoc-only an intentional constraint?

---

*Generated by automated static analysis. Because build/test/lint commands could not run in this environment, confirmation of the current pass/fail state and `npm audit` results is recommended before acting on dependency-level recommendations.*