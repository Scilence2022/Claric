/**
 * Imported skill store tests (localStorage CRUD, cap enforcement).
 */

/** localStorage mock (node test environment has no DOM). */
const localStorageMock = (() => {
    let store = {};
    return {
        getItem: (key) => (key in store ? store[key] : null),
        setItem: (key, value) => { store[key] = String(value); },
        removeItem: (key) => { delete store[key]; },
        clear: () => { store = {}; },
    };
})();
global.localStorage = localStorageMock;

const {
    loadImportedSkills,
    addImportedSkill,
    removeImportedSkill,
    __testing,
} = require('../src/lib/skill-store.js');

function makeDescriptor(name, extra = {}) {
    return {
        name,
        slash: `/${name}`,
        description: `${name} description`,
        category: 'chat',
        scope: 'chat',
        defaultTemplate: `template for ${name}`,
        imported: true,
        ...extra,
    };
}

describe('skill-store', () => {
    beforeEach(() => {
        localStorage.clear();
    });

    test('load returns an empty array when nothing is stored or corrupt', () => {
        expect(loadImportedSkills()).toEqual([]);
        localStorage.setItem('wordAI.skills.imported', '{not json');
        expect(loadImportedSkills()).toEqual([]);
    });

    test('add persists and round-trips a descriptor; re-import replaces by name', () => {
        expect(addImportedSkill(makeDescriptor('a')).ok).toBe(true);
        expect(addImportedSkill(makeDescriptor('b')).ok).toBe(true);
        expect(loadImportedSkills().map((s) => s.name)).toEqual(['a', 'b']);

        expect(addImportedSkill(makeDescriptor('a', { description: 'updated' })).ok).toBe(true);
        const skills = loadImportedSkills();
        expect(skills).toHaveLength(2);
        expect(skills.find((s) => s.name === 'a').description).toBe('updated');
    });

    test('enforces the import cap', () => {
        for (let i = 0; i < __testing.MAX_IMPORTED_SKILLS; i++) {
            expect(addImportedSkill(makeDescriptor(`skill-${i}`)).ok).toBe(true);
        }
        const result = addImportedSkill(makeDescriptor('one-too-many'));
        expect(result.ok).toBe(false);
        expect(result.error).toMatch(/too many/i);
        expect(loadImportedSkills()).toHaveLength(__testing.MAX_IMPORTED_SKILLS);
    });

    test('remove deletes by name', () => {
        addImportedSkill(makeDescriptor('a'));
        addImportedSkill(makeDescriptor('b'));
        removeImportedSkill('a');
        expect(loadImportedSkills().map((s) => s.name)).toEqual(['b']);
        // Unknown name is a no-op.
        removeImportedSkill('nope');
        expect(loadImportedSkills()).toHaveLength(1);
    });
});
