/**
 * Skill Store
 *
 * Persistence for imported skill packages (SKILL.md descriptors), kept
 * separate from PromptManager's prompt storage so the two sources can
 * evolve independently. Descriptors are stored under
 * `wordAI.skills.imported` as a JSON array; corrupt data fails closed to
 * an empty list.
 *
 * Pure module besides the guarded localStorage access — importable under
 * Jest/node.
 *
 * @module skill-store
 */

const STORAGE_KEY = 'wordAI.skills.imported';
/** Maximum number of imported skills kept. */
export const MAX_IMPORTED_SKILLS = 24;

function storage() {
    return typeof localStorage !== 'undefined' ? localStorage : null;
}

/**
 * Loads the imported skill descriptors. Always returns an array (empty on
 * missing/corrupt data).
 *
 * @returns {Array<object>}
 */
export function loadImportedSkills() {
    const store = storage();
    if (!store) return [];
    try {
        const parsed = JSON.parse(store.getItem(STORAGE_KEY));
        return Array.isArray(parsed) ? parsed : [];
    } catch (_err) {
        return [];
    }
}

/**
 * Persists the imported skill descriptors. Returns false when the write
 * fails (quota); the caller surfaces it.
 *
 * @param {Array<object>} skills
 * @returns {boolean}
 * @private
 */
function saveImportedSkills(skills) {
    const store = storage();
    if (!store) return false;
    try {
        store.setItem(STORAGE_KEY, JSON.stringify(skills));
        return true;
    } catch (_err) {
        return false;
    }
}

/**
 * Adds (or replaces, matching by name) one imported skill descriptor.
 *
 * @param {object} descriptor - Skill descriptor from parseSkillPackage
 * @returns {{ok: boolean, error?: string}}
 */
export function addImportedSkill(descriptor) {
    if (!descriptor || typeof descriptor !== 'object' || !descriptor.name) {
        return { ok: false, error: 'Invalid skill descriptor' };
    }
    const skills = loadImportedSkills();
    const filtered = skills.filter((s) => s && s.name !== descriptor.name);
    if (filtered.length >= MAX_IMPORTED_SKILLS) {
        return { ok: false, error: `Too many imported skills (max ${MAX_IMPORTED_SKILLS}). Remove one first.` };
    }
    filtered.push(descriptor);
    return saveImportedSkills(filtered) ? { ok: true } : { ok: false, error: 'Could not save the skill (storage quota)' };
}

/**
 * Removes one imported skill by name. Unknown names are a no-op.
 *
 * @param {string} name
 */
export function removeImportedSkill(name) {
    const skills = loadImportedSkills().filter((s) => s && s.name !== name);
    saveImportedSkills(skills);
}

/** @private */
export const __testing = { STORAGE_KEY, MAX_IMPORTED_SKILLS };
