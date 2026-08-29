/**
 * Settings tabs structural contract (markup-level regression guard).
 *
 * switchSettingsTab is generic: a `.settings-tab[data-tab="skills"]`
 * activates `#settingsPageSkills` by id convention. These assertions keep
 * the tab/page pair and the skill-import markup in sync — the import UI
 * moved here out of the Prompts page when it got its own tab.
 */

const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'src', 'taskpane', 'taskpane.html'), 'utf8');

describe('settings tabs markup contract', () => {
    test('a Skills tab exists and is paired with its page by id convention', () => {
        expect(html).toContain('data-tab="skills"');
        expect(html).toContain('id="settingsTabSkills"');
        expect(html).toContain('id="settingsPageSkills"');
        expect(html).toContain('aria-labelledby="settingsTabSkills"');
    });

    test('the skill-package import UI lives on the Skills page, not the Prompts page', () => {
        const skillsPage = html.slice(html.indexOf('id="settingsPageSkills"'));
        const promptsPage = html.slice(
            html.indexOf('id="settingsPagePrompts"'),
            html.indexOf('id="settingsPageSkills"'),
        );
        expect(skillsPage).toContain('id="skillImportBtn"');
        expect(skillsPage).toContain('id="skillImportList"');
        expect(promptsPage).not.toContain('skillImportBtn');
    });

    test('user-facing copy says slash commands, not skills, for the command list', () => {
        expect(html).toContain('Get started with these slash commands:');
        expect(html).toContain('Ask for help, or / for commands');
        expect(html).toContain('aria-label="Slash commands"');
        expect(html).not.toContain('Get started with these skills:');
        expect(html).not.toContain('/ for skills');
    });
});
