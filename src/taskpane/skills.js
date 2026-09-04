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
import { loadImportedSkills } from '../lib/skill-store.js';

/**
 * The built-in skills.
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
    Object.freeze({
        name: 'polish',
        slash: '/polish',
        description: 'Polish selected writing into clear, precise, formal academic English while preserving meaning',
        category: 'amendment',
        scope: 'selection-first',
        defaultTemplate:
            'Revise {selection} into polished academic English. Improve clarity, precision, flow, grammar, and scholarly ' +
            'tone while preserving the author\'s meaning, claims, citations, technical terminology, and level of certainty. ' +
            'Do not invent evidence or add substantive content. Keep paragraph structure unless a small change is necessary ' +
            'for coherence.',
    }),
    Object.freeze({
        name: 'simplify',
        slash: '/simplify',
        description: 'Make selected writing easier to understand without losing important meaning or necessary detail',
        category: 'amendment',
        scope: 'selection-first',
        defaultTemplate:
            'Simplify {selection} for a busy, intelligent reader. Use plain, direct language, shorter sentences, and clear ' +
            'transitions. Preserve all material facts, qualifications, numbers, citations, and intended meaning. Do not ' +
            'remove necessary technical terms; explain one only when the context requires it.',
    }),
    Object.freeze({
        name: 'shorten',
        slash: '/shorten',
        description: 'Make selected writing more concise while preserving its key message, evidence, and qualifications',
        category: 'amendment',
        scope: 'selection-first',
        defaultTemplate:
            'Shorten {selection} substantially without changing its meaning. Remove repetition, filler, and low-value ' +
            'wording; combine overlapping sentences and prefer economical phrasing. Preserve the key message, evidence, ' +
            'necessary context, qualifications, citations, and any required formatting or list structure.',
    }),
    Object.freeze({
        name: 'expand',
        slash: '/expand',
        description: 'Expand selected writing with useful explanation and transitions while avoiding unsupported claims',
        category: 'amendment',
        scope: 'selection-first',
        defaultTemplate:
            'Expand {selection} to make its reasoning and implications more complete. Add only useful explanation, ' +
            'definitions, examples, or transitions that follow from the existing text and context; do not invent facts, ' +
            'sources, results, or citations. Preserve the original position and level of certainty, and keep the writing ' +
            'focused and well organized.',
    }),
    Object.freeze({
        name: 'translate',
        slash: '/translate',
        description: 'Translate selected writing into a requested language while preserving meaning, terminology, and tone',
        category: 'amendment',
        scope: 'selection-first',
        defaultTemplate:
            'Translate {selection} into the target language specified by the user. Preserve the exact meaning, intent, ' +
            'technical terms, names, numbers, citations, formatting, and degree of certainty. Use natural, idiomatic ' +
            'language appropriate to the source text, but do not summarize, explain, or add content. If no target language ' +
            'is specified, translate English into Simplified Chinese and other languages into English.',
    }),
    Object.freeze({
        name: 'check-clarity',
        slash: '/check-clarity',
        description: 'Review writing for unclear wording, ambiguity, weak transitions, and unexplained references',
        category: 'comment',
        scope: 'selection-first',
        defaultTemplate:
            'Review {selection} for clarity and reader comprehension. Identify only genuine problems: ambiguous wording, ' +
            'unclear references, missing context, confusing sentence structure, weak transitions, or terminology that is ' +
            'not explained when explanation is needed. For each issue, add a concise comment quoting the exact text, ' +
            'explaining the concern, and suggesting a concrete revision. Do not praise, rewrite the text, or flag matters of ' +
            'mere preference. If there are no issues, reply with exactly: NO ISSUES FOUND.',
    }),
    Object.freeze({
        name: 'check-consistency',
        slash: '/check-consistency',
        description: 'Review the document for inconsistent terminology, facts, formatting, and internal conventions',
        category: 'comment',
        scope: 'document',
        defaultTemplate:
            'Review {selection} for internal consistency. Flag only conflicts or unexplained variations in defined terms, ' +
            'names, dates, numbers, units, capitalization, abbreviations, cross-references, headings, or formatting ' +
            'conventions. For each issue, quote the relevant text, identify the conflicting occurrence or rule, and suggest ' +
            'one concrete fix. Do not impose an external style guide or flag harmless variation. If the section is clean, ' +
            'reply with exactly: NO ISSUES FOUND.',
    }),
    Object.freeze({
        name: 'action-items',
        slash: '/action-items',
        description: 'Create an actionable list of decisions, tasks, owners, deadlines, and open questions from the document',
        category: 'summary',
        scope: 'document',
        defaultTemplate:
            'Extract actionable follow-up items from the document and any review comments. Produce a concise markdown table ' +
            'with columns: Action item, Owner, Deadline, Source or rationale, and Status. Include explicit and strongly ' +
            'implied tasks, decisions that need confirmation, and unresolved questions; mark unknown owners or deadlines as ' +
            'Not specified rather than guessing. Deduplicate overlapping items and preserve the document\'s terminology.\n\n' +
            'DOCUMENT:\n{whole document}\n\nREVIEW COMMENTS:\n{comments}',
    }),
    Object.freeze({
        name: 'executive-summary',
        slash: '/executive-summary',
        description: 'Generate a concise executive summary of the document, implications, risks, and decisions needed',
        category: 'summary',
        scope: 'document',
        defaultTemplate:
            'Write an executive summary for a busy decision-maker based on the document and review comments. Use markdown ' +
            'with these sections: Purpose, Key points, Decisions or recommendations, Risks and open issues, and Immediate ' +
            'next steps. Lead with the most important conclusions, distinguish facts from recommendations, and do not ' +
            'invent information. Keep it concise but specific, preserving material numbers, dates, owners, and caveats.\n\n' +
            'DOCUMENT:\n{whole document}\n\nREVIEW COMMENTS:\n{comments}',
    }),
    Object.freeze({
        name: 'key-points',
        slash: '/key-points',
        description: 'Extract the document\'s most important points in a concise, well-organized chat response',
        category: 'chat',
        scope: 'chat',
        defaultTemplate:
            'You identify the most important points in the current document for the user. Return a concise, prioritized ' +
            'markdown bullet list. Each point should state one substantive fact, obligation, finding, decision, or risk; ' +
            'include relevant numbers, dates, and qualifications, and group related points under short headings when helpful. ' +
            'Use the document context supplied to you, do not add unsupported interpretation, and follow any focus or audience ' +
            'the user specifies in the request.',
    }),
]);

/**
 * The reserved MCP-tools skill. Always listed; the turn runner explains
 * itself when no MCP server is configured yet (Settings → MCP Servers).
 * @type {Readonly<{name: string, slash: string, description: string, category: string, scope: string, defaultTemplate: string, reserved: boolean}>}
 */
export const RESERVED_MCP_SKILL = Object.freeze({
    name: 'mcp',
    slash: '/mcp',
    description: 'Run an instruction with tools from your configured MCP servers (Settings → MCP Servers)',
    category: 'tools',
    scope: 'tools',
    defaultTemplate: '',
    reserved: true,
});

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
 * Lists all available skills: the built-ins, the reserved MCP-tools
 * skill, one skill per prompt saved in PromptManager (slash name derived
 * from the prompt id), and the imported SKILL.md skill packages (from
 * skill-store).
 *
 * @param {object} promptManager - PromptManager instance
 * @returns {Array<object>} Skill descriptors (built-ins first)
 */
export function listSkills(promptManager) {
    const skills = [...BUILTIN_SKILLS, RESERVED_MCP_SKILL];
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

    // Imported SKILL.md packages (see lib/skill-package.js / lib/skill-store.js).
    skills.push(...loadImportedSkills());
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
