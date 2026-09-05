# Security Policy

## Supported versions

Security fixes are applied to the `main` branch. Release the fix as a patch
version and publish a new container image.

## Reporting a vulnerability

Please report vulnerabilities privately:

1. Open a GitHub Security Advisory (`Security` → `Report a vulnerability`), or
2. Email the maintainer via the address listed on the GitHub profile.

Include reproduction steps and, where possible, a minimal manifest/document
that triggers the issue. Please do not open public issues for suspected
security problems.

## Security posture

What this project does and does not promise:

- **LLM output sanitization** — generated summary markup is sanitized with
  DOMPurify before insertion into Word. This limits active markup; it does not
  prevent semantic prompt injection, incorrect edits, or external tool side effects.
- **Static server** — the production server (`scripts/docker-server.cjs`)
  rejects path traversal, malformed percent-encoded URLs, URLs containing
  control characters (a decoded NUL used to crash the process via
  `fs.readFile`), and non-GET/HEAD methods (except on the LLM proxy paths);
  it serves files under `dist/` plus the generated manifest. Unexpected
  throws are answered per-connection (500) instead of terminating the
  process.
- **LLM proxy** — production proxy routes are disabled by default; setting a
  provider's `*_PROXY_PATH` (e.g. `OLLAMA_PROXY_PATH=/ollama`) opts in and
  forwards request bodies and the `Authorization` header to the upstream set
  via the matching `*_PROXY_TARGET`. Configure only upstreams you trust.
- **TLS required** — Word requires HTTPS for add-in hosting. Certificate
  files are mounted read-only into the container.
- **API keys** — stored client-side in `localStorage` scoped to the add-in
  origin. They are sent only to the endpoint URL configured in Settings.
  Verify the endpoint URL before entering a key. Most providers use
  `Authorization: Bearer`; Claude uses `x-api-key`. A configured relay receives
  credentials before forwarding them. Claric does not encrypt browser storage
  or protect credentials against compromised same-origin scripts.
- **Dev server is dev-only** — the webpack dev server binds `127.0.0.1` by
  default (set `DEV_SERVER_HOST=0.0.0.0` to expose it). Its host allowlist defaults
  to loopback names; additional hosts must be explicitly configured and the
  special `all` value is filtered out. This is not authentication.
  The E2E/coding-agent endpoints (`/log`, `/api/e2e-loop/*`, `/api/test-cases`,
  `/api/prompts`) write bounded snapshots under the project root. They require
  a driver token and exact allowed Origin, and are registered only when
  `ENABLE_DEV_ENDPOINTS=true`. Treat the one-time startup token output as a secret.
  See [dev harness protocol](docs/dev-harness-protocol.md) for migration and
  storage limits. Never expose the dev server beyond your development machine.

## Known trade-offs

- The add-in sends document text (selections, comments, tracked changes) to
  the configured LLM backend. Deploy the backend under your control and over
  TLS; everything the LLM provider receives leaves the document.
- `.manifest-guid` is generated per deployment to keep Word's add-in identity
  stable; it is not a secret.
