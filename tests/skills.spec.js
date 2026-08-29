/**
 * Skill registry tests.
 *
 * Covers:
 *   - BUILTIN_SKILLS completeness (six frozen, well-formed descriptors)
 *   - resolveSkill parsing of "/name args" input
 *   - listSkills custom-skill registration from PromptManager state
 *
 * Pure JS -- no DOM or Word API required.
 */

const {
  BUILTIN_SKILLS,
  listSkills,
  resolveSkill,
} = require('../src/taskpane/skills.js');

const EXPECTED_SLASHES = [
  '/check-doc',
  '/copy-edit',
  '/summarize-contract',
  '/flag-issues',
  '/industry-overview',
  '/storylining',
];

describe('BUILTIN_SKILLS registry', () => {
  test('contains exactly the six built-in skills with unique slashes', () => {
    expect(BUILTIN_SKILLS.length).toBe(6);
    const slashes = BUILTIN_SKILLS.map((s) => s.slash);
    expect(new Set(slashes)).toEqual(new Set(EXPECTED_SLASHES));
  });

  test('registry array and entries are frozen', () => {
    expect(Object.isFrozen(BUILTIN_SKILLS)).toBe(true);
    BUILTIN_SKILLS.forEach((s) => expect(Object.isFrozen(s)).toBe(true));
  });

  test('every skill has name, slash, description, category, defaultTemplate, scope', () => {
    for (const skill of BUILTIN_SKILLS) {
      expect(typeof skill.name).toBe('string');
      expect(skill.slash).toBe(`/${skill.name}`);
      expect(skill.description.length).toBeGreaterThan(0);
      expect(['amendment', 'comment', 'summary', 'chat']).toContain(skill.category);
      expect(skill.defaultTemplate.length).toBeGreaterThan(0);
      expect(['selection-first', 'document', 'chat']).toContain(skill.scope);
    }
  });

  test('pipeline skills use a category matching their pipeline', () => {
    const byName = Object.fromEntries(BUILTIN_SKILLS.map((s) => [s.name, s]));
    expect(byName['check-doc'].category).toBe('comment');
    expect(byName['copy-edit'].category).toBe('amendment');
    expect(byName['summarize-contract'].category).toBe('summary');
    expect(byName['flag-issues'].category).toBe('comment');
    expect(byName['industry-overview'].category).toBe('chat');
    expect(byName['storylining'].category).toBe('chat');
  });
});

describe('resolveSkill', () => {
  test('parses slash command with args', () => {
    const result = resolveSkill('/copy-edit fix the typos please', BUILTIN_SKILLS);
    expect(result).not.toBeNull();
    expect(result.skill.name).toBe('copy-edit');
    expect(result.args).toBe('fix the typos please');
  });

  test('parses slash command without args', () => {
    const result = resolveSkill('/check-doc', BUILTIN_SKILLS);
    expect(result.skill.name).toBe('check-doc');
    expect(result.args).toBe('');
  });

  test('matches slash token case-insensitively', () => {
    const result = resolveSkill('/COPY-EDIT', BUILTIN_SKILLS);
    expect(result.skill.name).toBe('copy-edit');
  });

  test('returns null for non-slash input', () => {
    expect(resolveSkill('fix this text', BUILTIN_SKILLS)).toBeNull();
  });

  test('returns null for unknown slash command', () => {
    expect(resolveSkill('/unknown-skill do things', BUILTIN_SKILLS)).toBeNull();
  });

  test('returns null for non-string input', () => {
    expect(resolveSkill(null, BUILTIN_SKILLS)).toBeNull();
    expect(resolveSkill(undefined, BUILTIN_SKILLS)).toBeNull();
  });
});

describe('listSkills', () => {
  function makePromptManager(promptsByCategory) {
    return {
      getPrompts: (category) => promptsByCategory[category] || [],
    };
  }

  test('returns built-ins when promptManager is missing or empty', () => {
    expect(listSkills(null).length).toBe(6);
    const pm = makePromptManager({});
    expect(listSkills(pm).length).toBe(6);
  });

  test('registers saved prompts as custom slash commands', () => {
    const pm = makePromptManager({
      amendment: [{ id: 'legal-review', name: 'Legal Review', template: 'tpl', description: 'desc' }],
      summary: [{ id: 'exec-summary', name: 'Exec Summary', template: 'tpl', description: '' }],
    });
    const skills = listSkills(pm);
    expect(skills.length).toBe(8);

    const legal = skills.find((s) => s.slash === '/legal-review');
    expect(legal).toBeDefined();
    expect(legal.custom).toBe(true);
    expect(legal.category).toBe('amendment');
    expect(legal.scope).toBe('selection-first');
    expect(legal.defaultTemplate).toBe('tpl');

    const summary = skills.find((s) => s.slash === '/exec-summary');
    expect(summary.category).toBe('summary');
    expect(summary.scope).toBe('document');
  });

  test('custom skills are resolvable by resolveSkill', () => {
    const pm = makePromptManager({
      comment: [{ id: 'my-check', name: 'My Check', template: 'tpl', description: '' }],
    });
    const result = resolveSkill('/my-check extra', listSkills(pm));
    expect(result).not.toBeNull();
    expect(result.skill.custom).toBe(true);
    expect(result.args).toBe('extra');
  });

  test('context prompts map to chat scope', () => {
    const pm = makePromptManager({
      context: [{ id: 'persona', name: 'Persona', template: 'tpl', description: '' }],
    });
    const skill = listSkills(pm).find((s) => s.slash === '/persona');
    expect(skill.scope).toBe('chat');
    expect(skill.category).toBe('context');
  });
});

describe('imported SKILL.md packages in the registry', () => {
    const { addImportedSkill } = require('../src/lib/skill-store.js');

    // jsdom-free localStorage mock; skill-store reads the global lazily.
    beforeEach(() => {
        const store = {};
        global.localStorage = {
            getItem: (k) => (k in store ? store[k] : null),
            setItem: (k, v) => { store[k] = String(v); },
            removeItem: (k) => { delete store[k]; },
            clear: () => { for (const k of Object.keys(store)) delete store[k]; },
        };
    });

    afterEach(() => {
        delete global.localStorage;
    });

    function pm() {
        return { getPrompts: () => [] };
    }

    test('listSkills appends imported packages after builtins', () => {
        expect(addImportedSkill({
            name: 'literature-review', slash: '/literature-review',
            description: 'Review against the literature', category: 'chat',
            scope: 'chat', defaultTemplate: 'body', imported: true,
        }).ok).toBe(true);

        const skills = listSkills(pm());
        expect(skills[0].name).toBe('check-doc'); // builtins first
        const imported = skills.find((s) => s.name === 'literature-review');
        expect(imported).toEqual(expect.objectContaining({
            slash: '/literature-review', imported: true, category: 'chat',
        }));
        expect(resolveSkill('/literature-review focus on methods', skills).args).toBe('focus on methods');
    });

    test('an empty store yields no imported skills', () => {
        expect(listSkills(pm()).filter((s) => s.imported)).toHaveLength(0);
    });
});
