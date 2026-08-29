/**
 * Skill package (SKILL.md) tests.
 *
 * A skill package is the Claude-style format: YAML-ish frontmatter
 * (name, description, optional category/scope) over a markdown body that
 * becomes the prompt template. Parsing must be strict enough to reject
 * junk (no name/description, bad category) while tolerating unknown
 * frontmatter fields — skill packages evolve upstream.
 */

const {
    parseSkillPackage,
    buildSkillMarkdown,
} = require('../src/lib/skill-package.js');

const log = () => {};

describe('parseSkillPackage', () => {
    test('parses a minimal valid package', () => {
        const raw = [
            '---',
            'name: literature-review',
            'description: Review a draft against the literature and flag gaps',
            '---',
            'Read the draft and flag missing citations.',
            'Use {selection} for the current text.',
        ].join('\n');

        const pkg = parseSkillPackage(raw, log);

        expect(pkg).toEqual(expect.objectContaining({
            name: 'literature-review',
            slash: '/literature-review',
            description: 'Review a draft against the literature and flag gaps',
            category: 'chat',
            scope: 'chat',
            defaultTemplate: 'Read the draft and flag missing citations.\nUse {selection} for the current text.',
            imported: true,
        }));
    });

    test('honors optional category and scope from frontmatter', () => {
        const raw = [
            '---',
            'name: tighten-prose',
            'description: Tighten wordy prose as tracked changes',
            'category: amendment',
            'scope: selection-first',
            '---',
            'Tighten the following text without changing meaning.\n\n{selection}',
        ].join('\n');

        expect(parseSkillPackage(raw, log)).toEqual(expect.objectContaining({
            category: 'amendment',
            scope: 'selection-first',
        }));
    });

    test('derives the default scope from the category', () => {
        const raw = '---\nname: doc-audit\ndescription: d\ncategory: summary\n---\nbody';
        expect(parseSkillPackage(raw, log).scope).toBe('document');
    });

    test('rejects packages without frontmatter, name, description, or body', () => {
        expect(parseSkillPackage('no frontmatter at all', log)).toBeNull();
        expect(parseSkillPackage('---\ndescription: only desc\n---\nbody', log)).toBeNull();
        expect(parseSkillPackage('---\nname: x\n---\nbody without description', log)).toBeNull();
        expect(parseSkillPackage('---\nname: x\ndescription: d\n---\n   \n', log)).toBeNull();
    });

    test('rejects invalid names and unknown categories', () => {
        expect(parseSkillPackage('---\nname: Bad Name!\ndescription: d\n---\nbody', log)).toBeNull();
        expect(parseSkillPackage('---\nname: ok\ndescription: d\ncategory: nonsense\n---\nbody', log)).toBeNull();
        expect(parseSkillPackage('---\nname: ok\ndescription: d\nscope: sideways\n---\nbody', log)).toBeNull();
    });

    test('ignores unknown frontmatter fields (forward compatibility)', () => {
        const raw = [
            '---',
            'name: forward',
            'description: d',
            'license: MIT',
            'allowed-tools: web_search, calc',
            'version: 3',
            '---',
            'body text',
        ].join('\n');

        const pkg = parseSkillPackage(raw, log);
        expect(pkg.name).toBe('forward');
        expect(pkg.defaultTemplate).toBe('body text');
    });

    test('enforces the size cap on the package body', () => {
        const big = 'x'.repeat(40_000);
        const raw = `---\nname: too-big\ndescription: d\n---\n${big}`;
        expect(parseSkillPackage(raw, log)).toBeNull();
    });
});

describe('buildSkillMarkdown', () => {
    test('round-trips a descriptor back to markdown', () => {
        const raw = [
            '---',
            'name: tighten-prose',
            'description: Tighten wordy prose',
            'category: amendment',
            'scope: selection-first',
            '---',
            'Tighten the text.\n\n{selection}',
        ].join('\n');

        const pkg = parseSkillPackage(raw, log);
        const rebuilt = parseSkillPackage(buildSkillMarkdown(pkg), log);
        expect(rebuilt).toEqual(pkg);
    });

    test('omits optional fields when they equal the derived defaults', () => {
        const md = buildSkillMarkdown(parseSkillPackage('---\nname: chat-skill\ndescription: d\n---\nbody', log));
        expect(md).not.toContain('category:');
        expect(md).not.toContain('scope:');
        expect(md).toContain('name: chat-skill');
    });
});
