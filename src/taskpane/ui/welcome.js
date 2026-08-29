/**
 * Welcome Module
 *
 * Renders the empty-state welcome block with skill chip buttons. Clicking a
 * chip hands the skill's slash command to the input bar.
 *
 * @module ui/welcome
 */

/**
 * Renders skill chips into the welcome container.
 *
 * @param {Array<object>} skills - Skills to show (built-ins)
 * @param {function(object)} onPick - Called with the chosen skill
 */
export function renderWelcomeChips(skills, onPick) {
    const container = document.getElementById('skillChips');
    if (!container) return;

    container.innerHTML = '';
    for (const skill of skills) {
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = 'skill-chip';

        const name = document.createElement('span');
        name.className = 'skill-chip-name';
        name.textContent = skill.slash;
        const desc = document.createElement('span');
        desc.className = 'skill-chip-desc';
        desc.textContent = skill.description;

        chip.appendChild(name);
        chip.appendChild(desc);
        chip.addEventListener('click', () => onPick(skill));
        container.appendChild(chip);
    }
}


/**
 * Selects which skills earn a welcome-screen chip: the built-ins, the
 * reserved /mcp skill, and imported SKILL.md packages. Prompt-derived
 * customs are excluded — the user authored those and already knows them;
 * the welcome list stays a "get started" set.
 *
 * @param {Array<object>} skills - Full skill list (from listSkills)
 * @returns {Array<object>}
 */
export function selectWelcomeSkills(skills) {
    return (skills || []).filter((s) => s && !s.custom && !s.promptId);
}
