/**
 * Conversation Module
 *
 * Turn routing and orchestration for the chat-driven taskpane.
 *
 * Routing rules (routeTurn — pure, testable):
 *   - "/skill args"            -> skill turn (pipeline depends on skill.category)
 *   - free text + selection    -> selection edit (user text is the edit instruction)
 *   - free text, no selection  -> document Q&A (answer in chat)
 *
 * createConversation(deps) wires routing to the chat view, input bar, and
 * word actions. All Word/LLM side effects live behind deps.actions so the
 * orchestration is testable under Jest/node.
 *
 * @module conversation
 */

import * as defaultActions from './word-actions.js';
import { listSkills, resolveSkill } from './skills.js';
import { createProposalCard } from './ui/proposal-card.js';

/** Turn types emitted by routeTurn. */
export const TURN_TYPE = Object.freeze({
    SKILL: 'skill',
    SELECTION_EDIT: 'selection-edit',
    DOC_EDIT: 'doc-edit',
    DOC_QA: 'doc-qa',
});

/**
 * English edit-intent verbs (word-boundary matched) and Chinese edit-intent
 * substrings. Free text without a selection carrying one of these is an
 * instruction to edit the document, not a question about it.
 */
const EDIT_INTENT_RE = /\b(edit|revise|revision|polish|proofread|rewrite|redline|fix|amend|correct|improve|rephrase)\b|润色|修订|修改|批改|校对|校对|改写|审改|修正|完善/i;

/**
 * Leading question markers (EN + ZH). When the input STARTS with one of these
 * it is a question even if it contains an edit verb later ("how should I
 * improve this section?"), so Q&A takes precedence over edit intent.
 */
const QUESTION_LEAD_RE = /^\s*(what|why|how|does|do|is|are|can|could|should|would|which|who|when|where|explain|describe|summarize|list|tell me)\b|^\s*(什么|为什么|为何|怎么|怎样|如何|哪些|哪个|是不是|是否|能否|解释|说明|总结|概述|介绍)/i;

/**
 * True when free text (no selection) expresses an instruction to edit the
 * document rather than a question about it.
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeEditIntent(text) {
    if (QUESTION_LEAD_RE.test(text)) return false;
    return EDIT_INTENT_RE.test(text);
}

/**
 * Routes raw chat input to a turn descriptor. Pure function.
 *
 * @param {string} text - Raw chat input
 * @param {object} ctx
 * @param {boolean} ctx.hasSelection - Whether the document has a non-empty selection
 * @param {Array<object>} ctx.skills - Available skills (from listSkills)
 * @returns {{ type: string, skill?: object, args?: string, instruction?: string, question?: string } | null}
 *   Null for empty input.
 */
export function routeTurn(text, { hasSelection, skills } = {}) {
    const trimmed = (text || '').trim();
    if (!trimmed) return null;

    const resolved = resolveSkill(trimmed, skills);
    if (resolved) {
        return { type: TURN_TYPE.SKILL, skill: resolved.skill, args: resolved.args };
    }
    if (hasSelection) {
        return { type: TURN_TYPE.SELECTION_EDIT, instruction: trimmed };
    }
    if (looksLikeEditIntent(trimmed)) {
        return { type: TURN_TYPE.DOC_EDIT, instruction: trimmed };
    }
    return { type: TURN_TYPE.DOC_QA, question: trimmed };
}

/**
 * Creates the conversation controller.
 *
 * @param {object} deps
 * @param {object} deps.appState - Shared app state
 * @param {object} deps.view - Chat view API (addUserMessage, createAssistantMessage, addSystemNote, clearChat, renderWelcome)
 * @param {object} deps.input - Input bar API (setProcessing, setValue, focus)
 * @param {function} deps.log - Activity log callback
 * @param {function} [deps.logWithRetry] - Log-with-retry-link callback
 * @param {function} [deps.updateStatusBar] - Comment pending-count callback
 * @param {object} [deps.actions] - word-actions overrides (tests)
 * @param {function} [deps.getSelectionState] - async () => boolean (tests)
 * @returns {{ submit: Function, cancel: Function, newChat: Function }}
 */
export function createConversation(deps) {
    const { appState, view, input, log, logWithRetry, updateStatusBar } = deps;
    const actions = deps.actions || defaultActions;
    const getSelectionState = deps.getSelectionState || actions.hasNonEmptySelection;

    const actionDeps = { appState, log, logWithRetry, updateStatusBar };

    /**
     * True while any pipeline is running (blocks new turns).
     * @returns {boolean}
     */
    function isBusy() {
        return appState.isProcessing || appState.isProcessingDoc || appState.isProcessingSummary;
    }

    /**
     * Runs a document-scope skill with progress + citation pills.
     */
    async function runDocumentTurn(skill, args, msg) {
        appState.isProcessingDoc = true;
        appState.processDocController = new AbortController();
        input.setProcessing(true);
        try {
            msg.setStatus(`Processing document (${skill.name})...`);
            const commentInstructions = getCommentInstructions();
            const { applicationResult, chunks, cancelled } = await actions.runDocumentSkill(actionDeps, {
                category: skill.category,
                promptTemplate: withArgs(skill.defaultTemplate, args),
                commentInstructions,
                onProgress: (p) => msg.showProgress(p),
            });
            msg.hideProgress();
            if (cancelled) {
                msg.setStatus('Cancelled — already-applied changes remain in the document.');
            } else {
                msg.setStatus(
                    `Done: ${applicationResult.amendmentsApplied} amendment(s), ` +
                    `${applicationResult.commentsInserted} comment(s) across ${chunks.length} section(s).`
                );
                msg.addCitationPills(chunks.map(chunkCitation), (searchText) => {
                    actions.revealTextSnippet(actionDeps, searchText);
                });
            }
        } catch (error) {
            msg.hideProgress();
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled — already-applied changes remain in the document.');
            } else {
                msg.markError(`Document processing failed: ${error.message}`);
            }
        } finally {
            appState.isProcessingDoc = false;
            appState.processDocController = null;
            input.setProcessing(false);
        }
    }

    /**
     * Runs a selection-scope amendment turn and stages the proposal card.
     */
    async function runSelectionEditTurn(promptTemplate, msg) {
        appState.isProcessing = true;
        input.setProcessing(true);
        try {
            msg.setStatus('Drafting edit...');
            const commentInstructions = getCommentInstructions();
            const proposal = await actions.prepareSelectionAmendment(actionDeps, {
                promptTemplate,
                commentInstructions,
            });
            msg.setStatus('');
            const card = createProposalCard({
                title: 'Proposed edit',
                beforeChars: proposal.selectionText.length,
                afterChars: proposal.amendedText ? proposal.amendedText.length : 0,
                comment: proposal.commentText,
                onApply: async () => {
                    try {
                        await actions.applySelectionAmendment(actionDeps, proposal);
                        card.markApplied();
                    } catch (error) {
                        log(`Apply failed: ${error.message}`, 'error');
                        card.markError(error.message);
                    }
                },
                onReject: () => {
                    log('Proposal rejected by user.', 'info');
                },
            });
            msg.attachProposal(card.el);
        } catch (error) {
            msg.markError(error.message);
        } finally {
            appState.isProcessing = false;
            input.setProcessing(false);
        }
    }

    /**
     * Runs a selection-scope comment turn (fire-and-forget comment pipeline).
     */
    async function runSelectionCommentTurn(skill, args, msg) {
        if (!appState.supportsComments) {
            msg.markError('Comment features require Word API 1.4.');
            return;
        }
        try {
            await actions.fireSelectionComment(actionDeps, {
                promptTemplate: withArgs(skill.defaultTemplate, args),
            });
            msg.setStatus('Comment request fired — it will appear in the document shortly.');
        } catch (error) {
            msg.markError(error.message);
        }
    }

    /**
     * Runs the summary pipeline (result opens as a new document).
     */
    async function runSummaryTurn(skill, args, msg) {
        appState.isProcessingSummary = true;
        input.setProcessing(true);
        try {
            msg.setStatus('Generating summary document...');
            const result = await actions.runSummarySkill(actionDeps, {
                promptTemplate: withArgs(skill.defaultTemplate, args),
            });
            msg.setStatus(`Summary document created (${result.chars} chars${result.commentCount ? `, ${result.commentCount} comment(s) included` : ''}).`);
        } catch (error) {
            msg.markError(error.message);
        } finally {
            appState.isProcessingSummary = false;
            input.setProcessing(false);
        }
    }

    /**
     * Runs a chat Q&A turn with streaming.
     */
    async function runQaTurn(question, skillTemplate, msg) {
        appState.isProcessing = true;
        appState.chatController = new AbortController();
        input.setProcessing(true);
        try {
            msg.setStatus('Reading the document...');
            const answer = await actions.answerQuestion(actionDeps, {
                question,
                skillTemplate,
                signal: appState.chatController.signal,
                onStatus: (s) => msg.setStatus(s),
                onToken: (token) => {
                    msg.setStatus('');
                    msg.appendText(token);
                },
            });
            // Re-render with think tags stripped (tokens stream raw).
            msg.setText(answer);
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessing = false;
            appState.chatController = null;
            input.setProcessing(false);
        }
    }

    /**
     * Dispatches a skill turn by category and resolved scope.
     */
    async function runSkillTurn(skill, args, hasSelection, msg) {
        switch (skill.category) {
            case 'chat':
            case 'context':
                // Custom context prompts act as chat personas.
                await runQaTurn(args || skill.description, skill.defaultTemplate, msg);
                break;
            case 'summary':
                await runSummaryTurn(skill, args, msg);
                break;
            case 'comment':
                if (skill.scope === 'selection-first' && hasSelection) {
                    await runSelectionCommentTurn(skill, args, msg);
                } else {
                    await runDocumentTurn(skill, args, msg);
                }
                break;
            case 'amendment':
            default:
                if (skill.scope === 'selection-first' && hasSelection) {
                    await runSelectionEditTurn(withArgs(skill.defaultTemplate, args), msg);
                } else {
                    await runDocumentTurn(skill, args, msg);
                }
                break;
        }
    }

    /**
     * Submits raw chat input as a new turn.
     *
     * @param {string} text
     */
    async function submit(text) {
        const trimmed = (text || '').trim();
        if (!trimmed) return;

        if (isBusy()) {
            log('Already processing. Cancel the current run first.', 'warning');
            return;
        }

        let hasSelection = false;
        try {
            hasSelection = await getSelectionState();
        } catch (_err) {
            hasSelection = false;
        }

        const turn = routeTurn(trimmed, {
            hasSelection,
            skills: listSkills(appState.promptManager),
        });
        if (!turn) return;

        view.hideWelcome();
        view.addUserMessage(trimmed);
        input.setValue('');

        const msg = view.createAssistantMessage();

        try {
            if (turn.type === TURN_TYPE.SKILL) {
                await runSkillTurn(turn.skill, turn.args, hasSelection, msg);
            } else if (turn.type === TURN_TYPE.SELECTION_EDIT) {
                await runSelectionEditTurn(turn.instruction, msg);
            } else if (turn.type === TURN_TYPE.DOC_EDIT) {
                // Free-text edit instruction without a selection: run the
                // whole-document amendment pipeline with the user's text as
                // the edit template.
                await runDocumentTurn({
                    name: 'Edit', category: 'amendment', scope: 'document',
                    defaultTemplate: turn.instruction,
                }, undefined, msg);
            } else {
                await runQaTurn(turn.question, null, msg);
            }
        } catch (error) {
            msg.markError(error.message || String(error));
        }
    }

    /**
     * Cancels the in-flight run (document pipeline or chat stream).
     */
    function cancel() {
        if (appState.processDocController) {
            appState.processDocController.abort();
            log('Cancelling document processing...', 'warning');
        }
        if (appState.chatController) {
            appState.chatController.abort();
        }
    }

    /**
     * Clears the chat and returns to the welcome state.
     */
    function newChat() {
        if (isBusy()) {
            cancel();
        }
        view.clearChat();
        view.renderWelcome();
        input.setValue('');
        input.focus();
    }

    return { submit, cancel, newChat };
}

/**
 * Appends skill args (the text after "/name") to the template so they reach
 * the LLM as extra instructions.
 *
 * @param {string} template
 * @param {string} [args]
 * @returns {string}
 */
function withArgs(template, args) {
    if (!args) return template;
    return `${template}\n\nAdditional instructions from the user: ${args}`;
}

/**
 * Reads the optional comment instructions from the settings panel.
 * Returns '' when the element is absent (tests, non-DOM environments).
 *
 * @returns {string}
 */
function getCommentInstructions() {
    if (typeof document === 'undefined') return '';
    const field = document.getElementById('commentInstructions');
    return field ? field.value.trim() : '';
}

/**
 * Builds citation-pill data for a processed chunk: label and search text are
 * the chunk's first non-empty paragraph (heading or first ~6 words).
 *
 * @param {object} chunk - DocumentChunk ({ id, paragraphs })
 * @returns {{ label: string, searchText: string }}
 */
export function chunkCitation(chunk) {
    const firstPara = (chunk.paragraphs || []).map((p) => p.text || '').find((t) => t.trim()) || '';
    const words = firstPara.trim().split(/\s+/).filter(Boolean);
    const label = words.slice(0, 6).join(' ') || chunk.id || 'Section';
    return { label, searchText: firstPara.trim() };
}
