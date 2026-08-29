/** @jest-environment jsdom */

/**
 * Welcome screen tests: chip rendering and the get-started skill selector
 * (built-ins + reserved /mcp + imported packages; prompt customs excluded).
 */

const { renderWelcomeChips, selectWelcomeSkills } = require('../src/taskpane/ui/welcome.js');

function setupDom() {
    document.body.innerHTML = '<div id="skillChips"></div>';
}

describe('selectWelcomeSkills', () => {
    test('keeps builtins, the reserved /mcp skill, and imported packages; drops prompt customs', () => {
        const skills = [
            { name: 'check-doc', slash: '/check-doc' },
            { name: 'mcp', slash: '/mcp', reserved: true, category: 'tools' },
            { name: 'imported-skill', slash: '/imported-skill', imported: true },
            { name: 'my-prompt', slash: '/my-prompt', custom: true, promptId: 'my-prompt' },
        ];
        expect(selectWelcomeSkills(skills).map((s) => s.name)).toEqual([
            'check-doc', 'mcp', 'imported-skill',
        ]);
    });

    test('tolerates a missing list', () => {
        expect(selectWelcomeSkills(null)).toEqual([]);
    });
});

describe('renderWelcomeChips', () => {
    beforeEach(setupDom);

    test('renders one chip per skill and fills the input on click', () => {
        const picked = [];
        renderWelcomeChips([
            { slash: '/mcp', description: 'Run with MCP tools' },
            { slash: '/copy-edit', description: 'Fix errors' },
        ], (skill) => picked.push(skill.slash));

        const chips = document.querySelectorAll('.skill-chip');
        expect(chips).toHaveLength(2);
        expect(chips[0].textContent).toContain('/mcp');
        chips[0].click();
        expect(picked).toEqual(['/mcp']);
    });
});
