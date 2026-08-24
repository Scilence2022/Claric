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

- **LLM output sanitization** — markdown produced by the LLM is sanitized
  with DOMPurify before insertion into Word documents, mitigating
  prompt-injection turning into live markup in generated summaries.
- **Static server** — the production server (`scripts/docker-server.cjs`)
  rejects path traversal, malformed percent-encoded URLs, and non-GET/HEAD
  methods; it serves only files under `dist/` plus the generated manifest.
- **TLS required** — Word requires HTTPS for add-in hosting. Certificate
  files are mounted read-only into the container.
- **API keys** — stored client-side in `localStorage` scoped to the add-in
  origin. They are sent only to the endpoint URL configured in Settings.
  Verify the endpoint URL before entering a key: the add-in sends the key
  as an `Authorization: Bearer` header to whatever URL is configured.
- **Dev server is dev-only** — the webpack dev server binds `0.0.0.0`, allows
  all hosts, and serves E2E logs with permissive CORS. Never expose it
  beyond your development machine.

## Known trade-offs

- The add-in sends document text (selections, comments, tracked changes) to
  the configured LLM backend. Deploy the backend under your control and over
  TLS; everything the LLM provider receives leaves the document.
- `.manifest-guid` is generated per deployment to keep Word's add-in identity
  stable; it is not a secret.
