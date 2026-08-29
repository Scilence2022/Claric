/**
 * Skill Package
 *
 * Parser/serializer for the Claude-style SKILL.md skill package format:
 * YAML-ish frontmatter (name, description, optional category/scope) over a
 * markdown body that becomes the prompt template. This is the interchange
 * format for importing skills into the skill registry (and exporting them
 * back out).
 *
 * Pure module — no DOM, no localStorage, no Word API.
 *
 * @module skill-package
 */

import { MAX_SKILL_PACKAGE_CHARS } from './skill-limits.js';

const NAME_RE = /^[a-z0-9][a-z0-9-]*$/;
const SCOPES = ['selection-first', 'document', 'chat'];
// Skill categories (skills.js pipeline routing) — NOT PromptManager's
// prompt categories (context/amendment/comment/summary): a skill's
// 'chat' category means "answer in chat", which prompts don't have.
const SKILL_CATEGORIES = ['amendment', 'comment', 'summary', 'chat'];

/** Scope that matches a category when the package doesn't name one. */
function defaultScopeFor(category) {
    if (category === 'summary') return 'document';
    if (category === 'chat') return 'chat';
    return 'selection-first';
}

/**
 * Extracts and parses the frontmatter object from a SKILL.md document.
 * Tolerates `key: value` lines (with optional quoted values) and ignores
 * unknown keys — the upstream package format evolves.
 *
 * @param {string} text
 * @returns {{ fields: Record<string, string>, body: string } | null} Null when
 *   no frontmatter block exists
 * @private
 */
function splitFrontmatter(text) {
    if (!text.startsWith('---')) return null;
    const end = text.indexOf('\n---', 3);
    if (end === -1) return null;

    const block = text.slice(3, end).replace(/^\r?\n/, '');
    const body = text.slice(end + 4).replace(/^\r?\n/, '');
    /** @type {Record<string, string>} */
    const fields = {};
    for (const line of block.split(/\r?\n/)) {
        const idx = line.indexOf(':');
        if (idx <= 0) continue;
        const key = line.slice(0, idx).trim();
        let value = line.slice(idx + 1).trim();
        if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
            value = value.slice(1, -1);
        }
        if (key) fields[key] = value;
    }
    return { fields, body };
}

/**
 * Parses one SKILL.md document into a skill descriptor compatible with
 * skills.js. Returns null (with a logged reason) for anything unusable —
 * imports must degrade individually, never break the whole batch.
 *
 * @param {string} raw - Raw SKILL.md text
 * @param {function} [log] - Logging callback (message, type)
 * @returns {{name: string, slash: string, description: string, category: string,
 *   scope: string, defaultTemplate: string, imported: boolean} | null}
 */
export function parseSkillPackage(raw, log = () => {}) {
    const text = String(raw || '');
    if (text.length > MAX_SKILL_PACKAGE_CHARS) {
        log(`Skill package rejected: exceeds ${MAX_SKILL_PACKAGE_CHARS} chars`, 'warning');
        return null;
    }

    const parts = splitFrontmatter(text.trim());
    if (!parts) {
        log('Skill package rejected: missing --- frontmatter block', 'warning');
        return null;
    }

    const name = (parts.fields.name || '').trim().toLowerCase();
    if (!NAME_RE.test(name)) {
        log(`Skill package rejected: invalid name "${parts.fields.name || ''}" (lowercase letters, digits, hyphens)`, 'warning');
        return null;
    }
    const description = (parts.fields.description || '').trim();
    if (!description) {
        log(`Skill package "${name}" rejected: missing description`, 'warning');
        return null;
    }
    const body = parts.body.trim();
    if (!body) {
        log(`Skill package "${name}" rejected: empty instruction body`, 'warning');
        return null;
    }

    const category = (parts.fields.category || 'chat').trim();
    if (!SKILL_CATEGORIES.includes(category)) {
        log(`Skill package "${name}" rejected: unknown category "${category}"`, 'warning');
        return null;
    }
    const scope = (parts.fields.scope || defaultScopeFor(category)).trim();
    if (!SCOPES.includes(scope)) {
        log(`Skill package "${name}" rejected: unknown scope "${scope}"`, 'warning');
        return null;
    }

    return {
        name,
        slash: `/${name}`,
        description,
        category,
        scope,
        defaultTemplate: body,
        imported: true,
    };
}

/**
 * Serializes a skill descriptor back to the SKILL.md format. Optional
 * fields are omitted when they equal the value parsing would derive, so a
 * round-trip is byte-stable for minimal packages.
 *
 * @param {object} descriptor - Skill descriptor (from parseSkillPackage or skills.js)
 * @returns {string} SKILL.md text
 */
export function buildSkillMarkdown(descriptor) {
    const pkg = descriptor || {};
    const lines = ['---', `name: ${pkg.name}`, `description: ${pkg.description}`];
    if (pkg.category && pkg.category !== 'chat') lines.push(`category: ${pkg.category}`);
    if (pkg.scope && pkg.scope !== defaultScopeFor(pkg.category || 'chat')) lines.push(`scope: ${pkg.scope}`);
    lines.push('---', String(pkg.defaultTemplate || ''));
    return lines.join('\n');
}
