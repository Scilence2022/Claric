/**
 * Shared skill-package size limits (kept out of skill-package.js so the
 * Jest spec can import the cap without dragging prompt-manager's module
 * graph along).
 *
 * @module skill-limits
 */

/** Hard cap on one imported skill package (frontmatter + body), in chars. */
export const MAX_SKILL_PACKAGE_CHARS = 32_000;
