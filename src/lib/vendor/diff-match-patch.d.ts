/**
 * Type surface for the vendored diff-match-patch 1.0.5 (implementation:
 * ./diff-match-patch.js, a verbatim CJS copy — keep byte-identical).
 *
 * This sibling declaration shadows the .js file for the type checker, so the
 * third-party implementation stays out of checkJs. The API surface exposed to
 * src/lib/word-diff/* is dynamic (word-mode methods are monkey-patched onto
 * the prototype in diff-wordmode.js), so the default export is intentionally
 * `any`; call sites that need more precision cast locally.
 */
declare const DiffMatchPatch: any;

export default DiffMatchPatch;
