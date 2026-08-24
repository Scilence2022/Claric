/**
 * Skills Module
 *
 * Skill registry for the chat-driven taskpane. A skill is a declarative
 * descriptor that routes a chat turn to one of the existing pipelines:
 *
 *   category: 'amendment' -> tracked-changes edit pipeline
 *             'comment'   -> Word comment pipeline
 *             'summary'   -> summary-to-new-document pipeline
 *             'chat'      -> answer in chat using document context
 *
 *   scope: 'selection-first' -> run on the selection when one exists,
 *                               otherwise on the whole document
 *          'document'        -> always whole document
 *          'chat'            -> never touches the document for edits
 *
 * Built-in skills are frozen descriptors. Custom skills are derived from the
 * prompts saved in PromptManager (localStorage keys wordAI.prompts.*), so
 * user prompts automatically appear as slash commands.
 *
 * Pure module -- no DOM, no Word API. Safe to import under Jest/node.
 *
 * @module skills
 */

import { CATEGORIES } from '../lib/prompt-manager.js';

/**
 * The six built-in skills.
 * @type {ReadonlyArray<{name: string, slash: string, description: string, category: string, defaultTemplate: string, scope: string}>}
 */
export const BUILTIN_SKILLS = Object.freeze([
    Object.freeze({
        name: 'check-doc',
        slash: '/check-doc',
        description: 'Check the whole document for inconsistent defined terms, broken cross-references, and numbering gaps',
        category: 'comment',
        scope: 'document',
        defaultTemplate:
            'Review the following document section as a meticulous contract proofreader. Identify ONLY these issue types:\n' +
            '1. Defined terms used inconsistently (e.g. a term is defined with one capitalization or wording but used differently elsewhere).\n' +
            '2. Broken or stale cross-references (references to sections, clauses, schedules, or exhibits that do not exist or are mislabeled).\n' +
            '3. Numbering gaps or sequence errors (e.g. 1.1, 1.2, 1.4 with 1.3 missing; letters out of order).\n\n' +
            'If you find issues, write a concise comment listing each issue with the exact quoted text and the suggested fix. ' +
            'If the section is clean, reply with exactly: NO ISSUES FOUND.\n\n{selection}',
    }),
    Object.freeze({
        name: 'copy-edit',
        slash: '/copy-edit',
        description: 'Fix spelling, grammar, and punctuation as tracked changes (uses your selection if one exists)',
        category: 'amendment',
        scope: 'selection-first',
        defaultTemplate:
            'Copy-edit the following text. Fix ONLY spelling, grammar, punctuation, and obvious typographical errors. ' +
            'Do not rewrite for style, do not change meaning, do not add or remove content. ' +
            'Make the minimal set of corrections.\n\n{selection}',
    }),
    Object.freeze({
        name: 'summarize-contract',
        slash: '/summarize-contract',
        description: 'Generate a contract summary with key terms, parties, dates, and obligations in a new document',
        category: 'summary',
        scope: 'document',
        defaultTemplate:
            'Summarize the following contract. Produce a structured markdown summary with these sections:\n' +
            '## Overview (parties, purpose, effective date)\n' +
            '## Key Commercial Terms (pricing, payment, term, renewal)\n' +
            '## Main Obligations of Each Party\n' +
            '## Termination and Liability\n' +
            '## Notable or Unusual Provisions\n\n' +
            'Base the summary on the document text and any extracted review comments below.\n\n' +
            'DOCUMENT:\n{whole document}\n\nREVIEW COMMENTS:\n{comments}',
    }),
    Object.freeze({
        name: 'flag-issues',
        slash: '/flag-issues',
        description: 'Flag provisions that deviate from market-standard practice, with suggested fixes',
        category: 'comment',
        scope: 'document',
        defaultTemplate:
            'Review the following contract section as an experienced transactional lawyer. Flag any provision that ' +
            'deviates from market-standard practice for this type of agreement: one-sided indemnities, unusual liability ' +
            'caps, missing customary protections, atypical termination or renewal terms, ambiguous obligations.\n\n' +
            'For each flag, quote the exact provision text, explain why it is off-market, and suggest a concrete fix. ' +
            'If the section is market-standard, reply with exactly: NO ISSUES FOUND.\n\n{selection}',
    }),
    Object.freeze({
        name: 'industry-overview',
        slash: '/industry-overview',
        description: 'Answer in chat: industry background and market context relevant to this document',
        category: 'chat',
        scope: 'chat',
        defaultTemplate:
            'You are an industry analyst. Based on the document provided, give a concise overview of the relevant ' +
            'industry: market structure, key players, standard commercial practices, and current trends that bear on ' +
            'this document. Answer the user\'s question in that context.',
    }),
    Object.freeze({
        name: 'storylining',
        slash: '/storylining',
        description: 'Answer in chat: structure and storyline advice for this document (SCQA / pyramid principle)',
        category: 'chat',
        scope: 'chat',
        defaultTemplate:
            'You are an expert in structured communication (pyramid principle, SCQA). Based on the document provided, ' +
            'help the user sharpen its storyline: governing thought, key line, argument order, and gaps in the logic. ' +
            'Answer the user\'s question with concrete, actionable structure advice.',
    }),
]);

/**
 * Maps a PromptManager category to the skill scope a custom prompt should get.
 *
 * @param {string} category - One of CATEGORIES
 * @returns {string} 'selection-first' | 'document' | 'chat'
 */
function scopeForCategory(category) {
    if (category === 'summary') return 'document';
    if (category === 'context') return 'chat';
    return 'selection-first';
}

/**
 * Lists all available skills: the six built-ins plus one skill per prompt
 * saved in PromptManager (slash name derived from the prompt id).
 *
 * @param {object} promptManager - PromptManager instance
 * @returns {Array<object>} Skill descriptors (built-ins first)
 */
export function listSkills(promptManager) {
    const skills = [...BUILTIN_SKILLS];
    if (!promptManager) return skills;

    for (const category of CATEGORIES) {
        for (const prompt of promptManager.getPrompts(category)) {
            skills.push({
                name: prompt.name,
                slash: `/${prompt.id}`,
                description: prompt.description || `Custom ${category} prompt`,
                category,
                scope: scopeForCategory(category),
                defaultTemplate: prompt.template,
                custom: true,
                promptId: prompt.id,
                promptCategory: category,
            });
        }
    }
    return skills;
}

/**
 * Parses a chat input of the form "/name rest of the text..." into a skill
 * plus its argument string. Matching is on the exact slash token.
 *
 * @param {string} input - Raw chat input
 * @param {Array<object>} skills - Skill list (from listSkills)
 * @returns {{ skill: object, args: string } | null} Null when the input is not
 *   a slash command or matches no known skill
 */
export function resolveSkill(input, skills) {
    if (typeof input !== 'string') return null;
    const trimmed = input.trim();
    if (!trimmed.startsWith('/')) return null;

    const spaceIdx = trimmed.indexOf(' ');
    const token = (spaceIdx === -1 ? trimmed : trimmed.slice(0, spaceIdx)).toLowerCase();
    const args = spaceIdx === -1 ? '' : trimmed.slice(spaceIdx + 1).trim();

    const skill = (skills || []).find((s) => s.slash.toLowerCase() === token);
    if (!skill) return null;
    return { skill, args };
}
