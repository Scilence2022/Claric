/**
 * Settings tabs structural contract (markup-level regression guard).
 *
 * switchSettingsTab is generic: a `.settings-tab[data-tab="skills"]`
 * activates `#settingsPageSkills` by id convention. These assertions keep
 * the tab/page pairs and their section markup in sync — the import UI
 * moved out of the Prompts page when it got its own tab, and the MCP
 * server UI moved out of the General page the same way.
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

    test('an MCP Servers tab exists and is paired with its page by id convention', () => {
        expect(html).toContain('data-tab="mcp"');
        expect(html).toContain('id="settingsTabMcp"');
        expect(html).toContain('id="settingsPageMcp"');
        expect(html).toContain('aria-labelledby="settingsTabMcp"');
    });

    test('the MCP server UI lives on the MCP page, not the General page', () => {
        const generalPage = html.slice(
            html.indexOf('id="settingsPageGeneral"'),
            html.indexOf('id="settingsPagePrompts"'),
        );
        // The MCP page is the last page, right before the overlay closes;
        // slice from there to the end of the markup.
        const mcpPage = html.slice(html.indexOf('id="settingsPageMcp"'));
        expect(mcpPage).toContain('id="mcpServerAddBtn"');
        expect(mcpPage).toContain('id="mcpServerList"');
        expect(mcpPage).toContain('id="mcpStepBudget"');
        expect(generalPage).not.toContain('mcpServerAddBtn');
        expect(generalPage).not.toContain('mcpStepBudget');
    });

    test('the General page exposes thinking and temperature controls', () => {
        const generalPage = html.slice(
            html.indexOf('id="settingsPageGeneral"'),
            html.indexOf('id="settingsPagePrompts"'),
        );
        expect(generalPage).toContain('id="thinkingLevelSelect"');
        expect(generalPage).toContain('<option value="default">Default</option>');
        expect(generalPage).toContain('<option value="low">Low</option>');
        expect(generalPage).toContain('<option value="medium">Medium</option>');
        expect(generalPage).toContain('<option value="high">High</option>');
        // Model-aware explanations sit next to the two generation controls.
        expect(generalPage).toContain('id="thinkingLevelHint"');
        expect(generalPage).toContain('id="temperatureInput"');
        expect(generalPage).toContain('id="temperatureHint"');
        expect(generalPage).toContain('type="number"');
        expect(generalPage).toContain('min="0"');
        expect(generalPage).toContain('max="2"');
        expect(generalPage).toContain('step="0.1"');
        expect(generalPage).toContain('backends or models may ignore unsupported options');
    });

    test('the settings panel exposes an id so the drag/resize logic can bind to it', () => {
        expect(html).toContain('id="settingsPanel"');
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
