/**
 * Conversation Module
 *
 * Turn routing and orchestration for the chat-driven taskpane.
 *
 * Routing rules (routeTurn — pure, testable):
 *   - "/skill args"            -> skill turn (pipeline depends on skill.category)
 *   - free text hitting 2+ intent families -> compound turn (task planner decomposes
 *     into per-pipeline tasks, executed in order, one proposal card each)
 *   - free text + illustration intent -> design an SVG illustration, stage an image-insert proposal
 *   - free text + table intent -> create a native Word table (staged proposal; explicit
 *     dimensions resolve to an empty grid without an LLM call)
 *   - free text + append intent -> generate content, stage an append-to-end proposal
 *   - free text + format intent -> staged formatting/insert ops (existing text never rewritten)
 *   - free text + cleanup intent -> deterministic empty-paragraph deletion (staged, no LLM —
 *     the parser never sees blank paragraphs, so the text pipelines structurally cannot)
 *   - free text + selection    -> selection edit (user text is the edit instruction)
 *   - free text + edit intent  -> document amendment run (staged proposal)
 *   - free text + question lead -> document Q&A (answer in chat)
 *   - free text, zero intent hits -> compound turn (task planner classifies the
 *     ambiguous instruction into the right pipeline; planning failure falls back to Q&A)
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
import { describeFormatOp } from '../lib/format-ops.js';

/** Turn types emitted by routeTurn. */
export const TURN_TYPE = Object.freeze({
    SKILL: 'skill',
    SELECTION_EDIT: 'selection-edit',
    DOC_EDIT: 'doc-edit',
    DOC_APPEND: 'doc-append',
    FORMAT: 'format',
    TABLE: 'table',
    ILLUSTRATION: 'illustration',
    CLEANUP: 'cleanup',
    COMPOUND: 'compound',
    DOC_QA: 'doc-qa',
});

/**
 * English edit-intent verbs (word-boundary matched) and Chinese edit-intent
 * substrings. Free text without a selection carrying one of these is an
 * instruction to edit the document, not a question about it. Update/enrich
 * verbs (更新/丰富/充实/增补/扩写) require a document-ish object nearby, so
 * phrasings like "是谁更新的" stay out of the edit pipelines.
 */
const EDIT_INTENT_RE = /\b(edit|revise|revision|polish|proofread|rewrite|redline|fix|amend|correct|improve|rephrase)\b|\bupdate\b.{0,20}\b(document|doc|content|text)\b|润色|修订|修改|批改|校对|改写|审改|修正|完善|(更新|充实|增补|扩写).{0,4}(文档|文章|内容|正文|文本|文字|段落|章节|故事|论文|报告)/i;

/**
 * Leading question markers (EN + ZH). When the input STARTS with one of these
 * it is a question even if it contains an edit verb later ("how should I
 * improve this section?"), so Q&A takes precedence over edit intent.
 */
const QUESTION_LEAD_RE = /^\s*(what|why|how|does|do|is|are|can|could|should|would|which|who|when|where|explain|describe|summarize|list|tell me)\b|^\s*(什么|为什么|为何|怎么|怎样|如何|哪些|哪个|是不是|是否|能否|解释|说明|总结|概述|介绍)/i;

/**
 * True when free text starts with a question marker (EN + ZH).
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeQuestion(text) {
    return QUESTION_LEAD_RE.test(text);
}

/**
 * True when free text (no selection) expresses an instruction to edit the
 * document rather than a question about it.
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeEditIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return EDIT_INTENT_RE.test(text);
}

/**
 * Append-intent markers: the user wants NEW content generated and inserted
 * into the document (typically at the end), not an edit of existing text and
 * not a chat answer. Chinese substrings cover 追加/续写/写到文档-style
 * phrasings; English patterns require an explicit target ("...to the
 * document/end") so plain "add a paragraph" stays Q&A.
 */
const APPEND_INTENT_RE = /\bappend\b|\badd\b.{0,30}\bto the (document|doc|end)\b|\bcontinue writing\b|\bkeep writing\b|\bextend the (document|doc|text)\b|追加|续写|补写|接着写|继续写|写到文档|写入文档|加到文档|添加到文档|插入到文档|附加到文档|到文档末尾|到文末/i;

/**
 * Illustration-intent markers: the user wants NEW artwork designed and
 * inserted into the document. The dedicated nouns (插图/插画/配图/svg/...)
 * match directly; generic image words (图片/图像/image/picture) require a
 * creation verb nearby, so an edit that merely mentions an image
 * ("修改图像描述的措辞") still routes to an edit pipeline. These route to
 * the illustration pipeline (SVG design + image insertion).
 */
const ILLUSTRATION_INTENT_RE = /插图|插画|配图|扉页图|题图|头图|画作|绘制|\bsvg\b|\billustrat|\bartwork\b|(设计|生成|插入|添加|增加|创作|制作).{0,15}(图片|图像|图画)|(draw|generate|create|add|insert|design)\b.{0,20}\b(image|picture|drawing)\b/i;

/**
 * True when free text asks for an illustration to be designed and inserted.
 * Questions stay Q&A even when they mention artwork ("如何给文章配插图？").
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeIllustrationIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return ILLUSTRATION_INTENT_RE.test(text);
}

/**
 * True when free text asks for new content to be appended/inserted into the
 * document. Questions stay Q&A even when they mention appending
 * ("如何续写这篇文章？").
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeAppendIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return APPEND_INTENT_RE.test(text);
}

/**
 * Table-creation intent markers: the user wants a NEW native Word table.
 * Either a creation verb (插入/创建/生成/...; insert/create/add/...) sits near
 * the table noun, or explicit dimensions (N行N列 / 3x3) appear at all — a bare
 * "表格...行" mention is NOT enough, so edit instructions about an existing
 * table ("修改表格第二行") stay on the edit pipelines. Placement phrases like
 * 到文档末尾 name WHERE the table goes, which is why a table hit also
 * suppresses the append count in countIntentFamilies.
 */
const TABLE_INTENT_RE = /(插入|创建|新增|添加|生成|制作|绘制|画|做|建)[^。]{0,12}表格|\b(insert|create|add|make|generate|draw|build)\b.{0,20}\btable\b|[0-9零〇一二两三四五六七八九十]+\s*行\s*[0-9零〇一二两三四五六七八九十]+\s*列|\b\d+\s*[x×]\s*\d+\b.{0,12}\btable\b|\btable\b.{0,12}\b\d+\s*[x×]\s*\d+\b/i;

/**
 * True when free text asks for a new table to be created. Questions stay Q&A
 * even when they mention tables ("如何在文档中插入表格？").
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeTableIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return TABLE_INTENT_RE.test(text);
}

/**
 * Format-intent markers: the user wants formatting/styling changes (bold,
 * color, heading, alignment, ...) WITHOUT altering existing text. These
 * route to the format pipeline (JSON formatting ops), not the text-diff
 * amendment pipelines. The format pipeline also handles short structural
 * inserts (an insert op), so "增加文章标题" (add a title) correctly lands
 * here via 标题. Chinese substrings cover 加粗/标红/居中-style
 * phrasings; English terms are word-boundary matched.
 */
const FORMAT_INTENT_RE = /样式|格式|加粗|粗体|斜体|下划线|高亮|标红|字体|字号|颜色|居中|对齐|缩进|标题\s*[1-9一二三]?|设为标题|设置为标题|\bbold\b|\bitalic|underline|highlight|font\b|\bcolor\b|\bcenter(ed)?\b|\balign|indent|heading\s*[1-9]|format(ting)?\b/i;

/**
 * True when free text asks for formatting/styling changes only. Questions
 * stay Q&A even when they mention formatting ("如何修改样式？").
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeFormatIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return FORMAT_INTENT_RE.test(text);
}

/**
 * Cleanup-intent markers: the user wants redundant EMPTY paragraphs deleted
 * (删除多余的空段落 / 清除空行 / "delete empty paragraphs"). This cannot go
 * through the text pipelines — the document parser skips blank paragraphs
 * (so the LLM never sees them) and the reassembler excludes them from
 * alignment — so it routes to a deterministic Word.js cleanup instead.
 */
const CLEANUP_INTENT_RE = /(删除|清除|清理|去掉|移除|去除).{0,8}(空\s*段落|空白\s*段落|空行|空白行)|\b(delete|remove|clean\s*up|get rid of|strip)\b.{0,20}\b(empty|blank|whitespace)\b.{0,4}\b(paragraphs?|lines?)/i;

/**
 * True when free text asks to delete redundant empty paragraphs. Questions
 * stay Q&A even when they mention blank paragraphs ("为什么有多余的空段落？").
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeCleanupIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return CLEANUP_INTENT_RE.test(text);
}

/**
 * Counts how many intent families an instruction hits (illustration, table,
 * append, format, edit, cleanup). Questions count as zero — the looksLike*
 * guards keep them Q&A. A table hit suppresses the append count: placement
 * phrasing ("到文档末尾插入一个表格") names where the table goes, it is not a
 * second request for appended prose. A count of 2+ means a compound
 * instruction ("增加标题，并深度润色修改") that no single pipeline can serve;
 * those go through the task planner (TURN_TYPE.COMPOUND) instead of dropping
 * or refusing the non-matching parts.
 *
 * @param {string} text - Trimmed chat input
 * @returns {number}
 */
export function countIntentFamilies(text) {
    let count = 0;
    if (looksLikeIllustrationIntent(text)) count++;
    const table = looksLikeTableIntent(text);
    if (table) count++;
    if (!table && looksLikeAppendIntent(text)) count++;
    if (looksLikeFormatIntent(text)) count++;
    if (looksLikeEditIntent(text)) count++;
    if (looksLikeCleanupIntent(text)) count++;
    return count;
}

/**
 * Routes raw chat input to a turn descriptor. Pure function.
 *
 * @param {string} text - Raw chat input
 * @param {object} ctx
 * @param {boolean} ctx.hasSelection - Whether the document has a non-empty selection
 * @param {Array<object>} ctx.skills - Available skills (from listSkills)
 * @param {boolean} [ctx.allowCompound=true] - False disables the compound
 *   branch (planner fallback re-routes through single intent)
 * @returns {{ type: string, skill?: object, args?: string, instruction?: string, question?: string, scope?: string } | null}
 *   Null for empty input.
 */
export function routeTurn(text, { hasSelection, skills, allowCompound = true } = {}) {
    const trimmed = (text || '').trim();
    if (!trimmed) return null;

    const resolved = resolveSkill(trimmed, skills);
    if (resolved) {
        return { type: TURN_TYPE.SKILL, skill: resolved.skill, args: resolved.args };
    }
    // Compound instructions hit several intent families at once; the task
    // planner decomposes them into per-pipeline tasks, executed in order.
    if (allowCompound && countIntentFamilies(trimmed) >= 2) {
        return { type: TURN_TYPE.COMPOUND, instruction: trimmed };
    }
    // Illustration intent wins over the append/edit branches: the artifact
    // nouns (插图/配图/svg...) name an image to design and insert, not text
    // to append or existing text to edit.
    if (looksLikeIllustrationIntent(trimmed)) {
        return { type: TURN_TYPE.ILLUSTRATION, instruction: trimmed };
    }
    // Table intent wins over the append branch: 到文档末尾-style placement
    // phrases name where the new table goes, not a prose-append request. It
    // also wins over the selection branch: a table creation anchors on the
    // selection (before/after) rather than editing it.
    if (looksLikeTableIntent(trimmed)) {
        return { type: TURN_TYPE.TABLE, instruction: trimmed };
    }
    // Append intent wins over the selection/edit branches: the user explicitly
    // asked for new content in the document, not a rewrite of existing text.
    if (looksLikeAppendIntent(trimmed)) {
        return { type: TURN_TYPE.DOC_APPEND, instruction: trimmed };
    }
    // Format intent wins over the selection/edit branches too: formatting ops
    // never rewrite text, so they must not enter the text-diff pipelines.
    if (looksLikeFormatIntent(trimmed)) {
        return { type: TURN_TYPE.FORMAT, instruction: trimmed, scope: hasSelection ? 'selection' : 'document' };
    }
    // Cleanup intent is document-scope and deterministic: empty paragraphs
    // are invisible to the parser/LLM, so no text pipeline could serve this.
    if (looksLikeCleanupIntent(trimmed)) {
        return { type: TURN_TYPE.CLEANUP, instruction: trimmed };
    }
    if (hasSelection) {
        // Selection + a question is a question ABOUT the selection (answered
        // in chat with the selection as context), not an edit instruction.
        if (looksLikeQuestion(trimmed)) {
            return { type: TURN_TYPE.DOC_QA, question: trimmed };
        }
        return { type: TURN_TYPE.SELECTION_EDIT, instruction: trimmed };
    }
    if (looksLikeEditIntent(trimmed)) {
        return { type: TURN_TYPE.DOC_EDIT, instruction: trimmed };
    }
    // Clear questions skip the classifier: straight to Q&A, no extra call.
    if (looksLikeQuestion(trimmed)) {
        return { type: TURN_TYPE.DOC_QA, question: trimmed };
    }
    // Zero intent-family hits and not a question lead: the instruction is
    // ambiguous. Let the task planner classify it into the right pipeline
    // (a cheap call — the planner never sees document text). Planning
    // failure falls back to single-intent routing (Q&A) in runCompoundTurn.
    if (allowCompound) {
        return { type: TURN_TYPE.COMPOUND, instruction: trimmed };
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
 * @param {function} [deps.getSelectionText] - async () => string (current
 *   selection text, '' when empty; tests)
 * @returns {{ submit: Function, cancel: Function, newChat: Function }}
 */
export function createConversation(deps) {
    const { appState, view, input, log, logWithRetry, updateStatusBar } = deps;
    const actions = deps.actions || defaultActions;
    const getSelectionText = deps.getSelectionText || actions.readSelectionSnippet;
    // Optional dep fired whenever the live session has new content worth
    // persisting (after a turn settles, and before newChat wipes the array).
    // The bootstrap wires this to sessions.saveSession(...).
    const onTurnCommitted = typeof deps.onTurnCommitted === 'function' ? deps.onTurnCommitted : null;

    /**
     * Fires onTurnCommitted with the current session snapshot. Best-effort —
     * a persistence failure must not break the chat pipeline.
     */
    function _commitSession() {
        if (!onTurnCommitted) return;
        try {
            const session = (view && typeof view.getCurrentSession === 'function')
                ? view.getCurrentSession()
                : null;
            onTurnCommitted(session);
        } catch (_err) {
            // Persistence errors must never break the live turn.
        }
    }

    /**
     * Builds per-turn action deps whose log lines also stream into the
     * assistant message's collapsible work log (Claude Code style).
     *
     * @param {object} msg - Assistant message handle
     * @returns {object} action deps for one turn
     */
    function actionDepsFor(msg) {
        return {
            appState,
            logWithRetry,
            updateStatusBar,
            log: (message, type) => {
                log(message, type);
                msg.appendLogLine(message);
            },
        };
    }

    /**
     * True while any pipeline is running (blocks new turns).
     * @returns {boolean}
     */
    function isBusy() {
        return appState.isProcessing || appState.isProcessingDoc || appState.isProcessingSummary;
    }

    /**
     * Runs a document-scope skill with progress + citation pills.
     * Amendment runs are gated: the LLM results are staged first and only
     * written to the document when the user applies the proposal card.
     */
    async function runDocumentTurn(skill, args, msg, turnDeps) {
        const gated = skill.category === 'amendment';
        const myController = new AbortController();
        appState.isProcessingDoc = true;
        appState.processDocController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus(`Processing document (${skill.name})...`);
            const commentInstructions = getCommentInstructions();
            const outcome = await actions.runDocumentSkill(turnDeps, {
                category: skill.category,
                promptTemplate: withArgs(skill.defaultTemplate, args),
                commentInstructions,
                signal: myController.signal,
                onProgress: (p) => msg.showProgress(p),
                onChunkToken: (info, kind, token) => msg.appendModelToken(info, kind, token),
                gateApply: gated,
            });
            msg.hideProgress();

            if (outcome.staged) {
                await stageDocumentProposal(outcome, msg, turnDeps);
                return;
            }

            const { applicationResult, chunks, cancelled } = outcome;
            if (cancelled) {
                msg.setStatus('Cancelled — already-applied changes remain in the document.');
            } else {
                msg.setStatus(
                    `Done: ${applicationResult.amendmentsApplied} amendment(s), ` +
                    `${applicationResult.commentsInserted} comment(s) across ${chunks.length} section(s).`
                );
                msg.addCitationPills(chunks.map(chunkCitation), (searchText) => {
                    actions.revealTextSnippet(turnDeps, searchText);
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
            // Only release state if we're still the active turn: cancel()
            // may have already nulled the controller to free the UI, and an
            // already-cancelled orphan whose background work settles late
            // must NOT clobber a follow-up turn's processing flags.
            if (appState.processDocController === myController) {
                appState.isProcessingDoc = false;
                appState.processDocController = null;
                input.setProcessing(false);
            }
        }
    }

    /**
     * Renders the proposal card for a staged document amendment run.
     * Apply writes the staged results as tracked changes; Reject discards
     * them (word-actions cleans up the chunk bookmarks either way).
     */
    async function stageDocumentProposal(outcome, msg, turnDeps) {
        // Only offer chunks whose amendment actually differs from the
        // original text — an LLM echo of the input is not a proposal.
        const amendedChunks = outcome.results.filter((r) => r.status === 'fulfilled'
            && r.amendment
            && _normalizeText(r.amendment) !== _normalizeText(chunkOriginalText(r)));

        if (amendedChunks.length === 0) {
            await outcome.discard();
            msg.setStatus(outcome.cancelledCount > 0
                ? 'Cancelled — no changes were applied.'
                : 'The model proposed no changes.');
            return;
        }

        const beforeChars = amendedChunks.reduce((s, r) => s + chunkOriginalText(r).length, 0);
        const afterChars = amendedChunks.reduce((s, r) => s + (r.amendment || '').length, 0);
        msg.setStatus('');

        const card = createProposalCard({
            title: `Proposed edits to ${amendedChunks.length} section(s)`,
            beforeChars,
            afterChars,
            comment: null,
            // One selectable change per amended section: inline diff plus a
            // locate-in-document link; Apply writes only the checked sections.
            items: amendedChunks.map((r) => {
                const citation = chunkCitation(r.chunk);
                return {
                    id: r.chunk.id,
                    label: citation.label,
                    before: chunkOriginalText(r),
                    after: r.amendment,
                    searchText: citation.searchText,
                };
            }),
            onLocate: (text) => actions.revealTextSnippet(turnDeps, text),
            onApply: async (selectedChunkIds) => {
                appState.isProcessingDoc = true;
                input.setProcessing(true);
                try {
                    const applicationResult = await outcome.apply(selectedChunkIds);
                    const applyErrors = applicationResult.errors || [];
                    for (const applyError of applyErrors) log(`Apply: ${applyError}`, 'warning');
                    if (applicationResult.amendmentsApplied === 0) {
                        // Honest terminal state: settling on "Applied" would be
                        // a lie when nothing (or only comments) landed.
                        card.markWarning(applyErrors.length
                            ? `Nothing applied: ${applyErrors[0]}`
                            : 'Nothing applied — the staged edits already match the document.');
                        msg.setStatus('');
                        return;
                    }
                    card.markApplied();
                    msg.setStatus(
                        `Done: ${applicationResult.amendmentsApplied} amendment(s), ` +
                        `${applicationResult.commentsInserted} comment(s) across ${outcome.chunks.length} section(s).` +
                        (applyErrors.length ? ` ${applyErrors.length} section(s) skipped — see activity log.` : '')
                    );
                    msg.addCitationPills(outcome.chunks.map(chunkCitation), (searchText) => {
                        actions.revealTextSnippet(turnDeps, searchText);
                    });
                } catch (error) {
                    log(`Apply failed: ${error.message}`, 'error');
                    card.markError(error.message);
                } finally {
                    appState.isProcessingDoc = false;
                    input.setProcessing(false);
                }
            },
            onReject: async () => {
                try {
                    await outcome.discard();
                } catch (error) {
                    log(`Discard failed: ${error.message}`, 'warning');
                }
                msg.setStatus('Proposed edits discarded.');
            },
        });
        msg.attachProposal(card, {
            title: `Proposed edits to ${amendedChunks.length} section(s)`,
            state: 'pending',
            countsText: `${beforeChars} → ${afterChars} chars`,
            items: amendedChunks.map((r) => {
                const citation = chunkCitation(r.chunk);
                return {
                    id: r.chunk.id,
                    label: citation.label,
                    before: chunkOriginalText(r),
                    after: r.amendment,
                    searchText: citation.searchText,
                };
            }),
        });
    }

    /**
     * Runs a selection-scope amendment turn and stages the proposal card.
     */
    async function runSelectionEditTurn(promptTemplate, msg, turnDeps, turnController) {
        const myController = turnController || new AbortController();
        appState.isProcessing = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Drafting edit...');
            const commentInstructions = getCommentInstructions();
            const proposal = await actions.prepareSelectionAmendment(turnDeps, {
                promptTemplate,
                commentInstructions,
                signal: myController.signal,
                onToken: (t) => msg.appendModelToken({ id: 'selection' }, 'content', t),
                onReasoning: (t) => msg.appendModelToken({ id: 'selection' }, 'reasoning', t),
            });
            msg.setStatus('');

            // Table patches review per-cell / per-row items instead of a
            // single before/after text diff.
            const isTable = !!proposal.tablePatch;
            if (isTable && proposal.tableItems.length === 0) {
                msg.setStatus('The model proposed no changes.');
                return;
            }
            const title = isTable ? 'Proposed table edit' : 'Proposed edit';
            const beforeChars = isTable
                ? proposal.tableItems.reduce((sum, item) => sum + (item.before || '').length, 0)
                : proposal.selectionText.length;
            const afterChars = isTable
                ? proposal.tableItems.reduce((sum, item) => sum + (item.after || '').length, 0)
                : (proposal.amendedText ? proposal.amendedText.length : 0);
            const items = isTable
                ? proposal.tableItems.map((item, index) => ({
                    id: index,
                    label: item.label,
                    before: item.before,
                    after: item.after,
                    searchText: item.searchText,
                }))
                : (proposal.amendedText ? [{
                    id: 'selection',
                    label: 'Selection rewrite',
                    before: proposal.selectionText,
                    after: proposal.amendedText,
                    searchText: proposal.selectionText.trim().slice(0, 60),
                }] : undefined);

            const card = createProposalCard({
                title,
                beforeChars,
                afterChars,
                comment: proposal.commentText,
                items,
                onLocate: (text) => actions.revealTextSnippet(turnDeps, text),
                onApply: async (selectedIds) => {
                    try {
                        // Table cards list one checkbox per cell patch, then
                        // one per row op — honor the user's unchecking.
                        let toApply = proposal;
                        if (isTable && selectedIds) {
                            const cellCount = proposal.tablePatch.cells.length;
                            const picked = new Set(selectedIds);
                            toApply = {
                                ...proposal,
                                tablePatch: {
                                    ...proposal.tablePatch,
                                    cells: proposal.tablePatch.cells.filter((_, i) => picked.has(i)),
                                    rowOps: proposal.tablePatch.rowOps.filter((_, j) => picked.has(cellCount + j)),
                                },
                            };
                        }
                        await actions.applySelectionAmendment(turnDeps, toApply);
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
            msg.attachProposal(card, {
                title,
                state: 'pending',
                countsText: `${beforeChars} → ${afterChars} chars`,
                items: items || [],
            });
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessing = false;
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Runs an append turn: the LLM drafts new content against the document
     * context, staged in a proposal card. Apply inserts it at the document
     * end as tracked changes; nothing is written before that.
     */
    async function runAppendTurn(instruction, msg, turnDeps, selectionText, turnController) {
        const myController = turnController || new AbortController();
        appState.isProcessing = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Drafting content to append...');
            const proposal = await actions.prepareDocumentAppend(turnDeps, {
                instruction,
                selectionText,
                signal: myController.signal,
                onToken: (t) => msg.appendModelToken({ id: 'append' }, 'content', t),
                onReasoning: (t) => msg.appendModelToken({ id: 'append' }, 'reasoning', t),
            });
            if (!proposal.generatedText) {
                msg.setStatus('The model returned no content to append.');
                return;
            }
            msg.setStatus('');
            const card = createProposalCard({
                title: 'Proposed content to append at the document end',
                beforeChars: 0,
                afterChars: proposal.generatedText.length,
                comment: null,
                onApply: async () => {
                    try {
                        await actions.applyDocumentAppend(turnDeps, proposal);
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
            msg.attachProposal(card, {
                title: 'Proposed content to append at the document end',
                state: 'pending',
                countsText: `0 → ${proposal.generatedText.length} chars`,
                items: [],
            });
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessing = false;
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Runs a format turn: the LLM emits JSON ops against the selection or
     * document scope — font/paragraph changes plus short structural inserts
     * (e.g. a title) — staged in a proposal card. Apply writes them via
     * Word.js (tracked when track-changes is on); existing text content is
     * never rewritten by this pipeline.
     */
    async function runFormatTurn(turn, msg, turnDeps, selectionText, turnController) {
        const myController = turnController || new AbortController();
        appState.isProcessing = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Planning document changes...');
            const proposal = await actions.prepareFormatProposal(turnDeps, {
                instruction: turn.instruction,
                scope: turn.scope,
                selectionText,
                signal: myController.signal,
                onToken: (t) => msg.appendModelToken({ id: 'format' }, 'content', t),
                onReasoning: (t) => msg.appendModelToken({ id: 'format' }, 'reasoning', t),
            });
            if (!proposal.ops || proposal.ops.length === 0) {
                msg.setStatus('The model proposed no changes.');
                return;
            }
            msg.setStatus('');
            const card = createProposalCard({
                title: `Proposed changes (${turn.scope} scope)`,
                countsText: `${proposal.ops.length} change op(s)`,
                comment: null,
                // One selectable change per op; Apply runs only the checked ops.
                items: proposal.ops.map((op, index) => ({
                    id: index,
                    label: describeFormatOp(op),
                    searchText: op.match ? op.match.trim().slice(0, 60) : undefined,
                })),
                onLocate: (text) => actions.revealTextSnippet(turnDeps, text),
                onApply: async (selectedIds) => {
                    try {
                        const ops = proposal.ops.filter((_, index) => selectedIds.includes(index));
                        const fmtResult = await actions.applyFormatProposal(turnDeps, { ...proposal, ops });
                        if (fmtResult && fmtResult.appliedRanges === 0 && fmtResult.insertedParagraphs === 0) {
                            card.markWarning('Nothing applied — no formatting targets matched. See the activity log.');
                        } else {
                            card.markApplied();
                        }
                    } catch (error) {
                        log(`Apply failed: ${error.message}`, 'error');
                        card.markError(error.message);
                    }
                },
                onReject: () => {
                    log('Proposal rejected by user.', 'info');
                },
            });
            msg.attachProposal(card, {
                title: `Proposed changes (${turn.scope} scope)`,
                state: 'pending',
                countsText: `${proposal.ops.length} change op(s)`,
                items: proposal.ops.map((op, index) => ({
                    id: index,
                    label: describeFormatOp(op),
                    searchText: op.match ? op.match.trim().slice(0, 60) : undefined,
                })),
            });
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessing = false;
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Runs a table turn: creates one NEW native Word table from a natural-
     * language request. Explicit dimensions without content wording resolve
     * to an empty grid deterministically (no LLM); content-bearing or
     * dimensionless requests go through the model's strict JSON table
     * protocol (lib/table-ops.js). The spec is staged in a proposal card
     * with a read-only grid preview; Apply inserts the table via Word.js
     * (tracked only where the host records structural revisions).
     */
    async function runTableTurn(turn, msg, turnDeps, turnController) {
        if (appState.supportsTables === false) {
            msg.markError('This Word host does not support the table APIs (WordApi 1.3). Update Word to create tables.');
            return;
        }
        const myController = turnController || new AbortController();
        appState.isProcessing = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Drafting table...');
            const proposal = await actions.prepareTableProposal(turnDeps, {
                instruction: turn.instruction,
                signal: myController.signal,
                onToken: (t) => msg.appendModelToken({ id: 'table' }, 'content', t),
                onReasoning: (t) => msg.appendModelToken({ id: 'table' }, 'reasoning', t),
            });
            if (!proposal.spec) {
                msg.setStatus('No table could be drafted from this instruction.');
                return;
            }
            msg.setStatus('');
            const rowCount = proposal.spec.rows.length;
            const columnCount = proposal.spec.rows[0].length;
            const title = `Proposed table (${rowCount}×${columnCount}, ${proposal.spec.position})`;
            const countsText = `${rowCount}×${columnCount} table at ${proposal.spec.position}`;
            const tablePreview = {
                rows: proposal.spec.rows,
                headerRowCount: proposal.spec.headerRowCount,
                style: proposal.spec.style,
                position: proposal.spec.position,
            };
            const card = createProposalCard({
                title,
                countsText,
                tablePreview,
                comment: null,
                onApply: async () => {
                    try {
                        const result = await actions.applyTableProposal(turnDeps, proposal);
                        if (result && result.warnings && result.warnings.length > 0) {
                            card.markWarning(`Table inserted with warning: ${result.warnings[0]}`);
                        } else {
                            card.markApplied();
                        }
                    } catch (error) {
                        log(`Apply failed: ${error.message}`, 'error');
                        card.markError(error.message);
                    }
                },
                onReject: () => {
                    log('Proposal rejected by user.', 'info');
                },
            });
            msg.attachProposal(card, {
                title,
                state: 'pending',
                countsText,
                tablePreview,
                items: [],
            });
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessing = false;
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Runs an illustration turn: the LLM designs one self-contained SVG from
     * the document's subject and mood, staged in a proposal card with an
     * image preview. Apply rasterizes it to PNG and inserts it into the
     * document via Word.js (tracked when track-changes is on).
     */
    async function runIllustrationTurn(turn, msg, turnDeps, turnController) {
        const myController = turnController || new AbortController();
        appState.isProcessing = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Designing illustration...');
            const proposal = await actions.prepareIllustrationProposal(turnDeps, {
                instruction: turn.instruction,
                signal: myController.signal,
                onToken: (t) => msg.appendModelToken({ id: 'illustration' }, 'content', t),
                onReasoning: (t) => msg.appendModelToken({ id: 'illustration' }, 'reasoning', t),
            });
            if (!proposal.svg) {
                msg.setStatus('The model produced no usable SVG illustration.');
                return;
            }
            msg.setStatus('');
            const positionLabel = proposal.positionLabel || `document ${proposal.position}`;
            const card = createProposalCard({
                title: 'Proposed illustration',
                countsText: `SVG ${(proposal.svg.length / 1024).toFixed(1)} KB → PNG at ${positionLabel}`,
                previewSrc: `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(proposal.svg)))}`,
                comment: null,
                onApply: async () => {
                    try {
                        await actions.applyIllustrationProposal(turnDeps, proposal);
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
            msg.attachProposal(card, {
                title: 'Proposed illustration',
                state: 'pending',
                countsText: `SVG ${(proposal.svg.length / 1024).toFixed(1)} KB → PNG at ${positionLabel}`,
                previewSrc: `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(proposal.svg)))}`,
                items: [],
            });
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessing = false;
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Runs a cleanup turn: deterministically scans the document for redundant
     * empty paragraphs (no LLM — blank paragraphs are invisible to the parser
     * and excluded from text-pipeline alignment, so a Word.js scan is the only
     * way to serve this) and stages a proposal card. Apply deletes them as
     * tracked changes.
     */
    async function runCleanupTurn(msg, turnDeps) {
        appState.isProcessing = true;
        input.setProcessing(true);
        try {
            msg.setStatus('Scanning for empty paragraphs...');
            const proposal = await actions.prepareEmptyParagraphCleanup(turnDeps);
            if (!proposal.emptyCount) {
                msg.setStatus('No empty paragraphs found.');
                return;
            }
            msg.setStatus('');
            const card = createProposalCard({
                title: 'Proposed cleanup',
                countsText: `Delete ${proposal.emptyCount} empty paragraph(s)`,
                comment: null,
                onApply: async () => {
                    try {
                        const result = await actions.applyEmptyParagraphCleanup(turnDeps);
                        if (!result || result.deleted === 0) {
                            card.markWarning('Nothing applied — no empty paragraphs remained.');
                        } else {
                            card.markApplied();
                        }
                    } catch (error) {
                        log(`Apply failed: ${error.message}`, 'error');
                        card.markError(error.message);
                    }
                },
                onReject: () => {
                    log('Proposal rejected by user.', 'info');
                },
            });
            msg.attachProposal(card, {
                title: 'Proposed cleanup',
                state: 'pending',
                countsText: `Delete ${proposal.emptyCount} empty paragraph(s)`,
                items: [],
            });
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
    async function runSelectionCommentTurn(skill, args, msg, turnDeps) {
        if (!appState.supportsComments) {
            msg.markError('Comment features require Word API 1.4.');
            return;
        }
        try {
            await actions.fireSelectionComment(turnDeps, {
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
    async function runSummaryTurn(skill, args, msg, turnDeps, turnController) {
        const myController = turnController || new AbortController();
        appState.isProcessingSummary = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Generating summary document...');
            const result = await actions.runSummarySkill(turnDeps, {
                promptTemplate: withArgs(skill.defaultTemplate, args),
                signal: myController.signal,
                onToken: (t) => {
                    msg.setStatus('');
                    msg.appendText(t);
                },
                onReasoning: (t) => msg.appendModelToken({ id: 'summary' }, 'reasoning', t),
            });
            msg.setStatus(`Summary document created (${result.chars} chars${result.commentCount ? `, ${result.commentCount} comment(s) included` : ''}).`);
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessingSummary = false;
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Runs a chat Q&A turn with streaming.
     * selectionText (when non-empty) is added to the prompt as a focused
     * excerpt alongside the full document context.
     */
    async function runQaTurn(question, skillTemplate, msg, turnDeps, selectionText, turnController) {
        const myController = turnController || new AbortController();
        appState.isProcessing = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Reading the document...');
            const answer = await actions.answerQuestion(turnDeps, {
                question,
                skillTemplate,
                selectionText,
                signal: myController.signal,
                onStatus: (s) => msg.setStatus(s),
                onToken: (token) => {
                    msg.setStatus('');
                    msg.appendText(token);
                },
                onReasoning: (t) => msg.appendModelToken({ id: 'qa' }, 'reasoning', t),
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
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Dispatches a skill turn by category and resolved scope.
     */
    async function runSkillTurn(skill, args, hasSelection, msg, turnDeps, selectionText, turnController) {
        switch (skill.category) {
            case 'chat':
            case 'context':
                // Custom context prompts act as chat personas.
                await runQaTurn(args || skill.description, skill.defaultTemplate, msg, turnDeps, selectionText, turnController);
                break;
            case 'summary':
                await runSummaryTurn(skill, args, msg, turnDeps, turnController);
                break;
            case 'comment':
                if (skill.scope === 'selection-first' && hasSelection) {
                    await runSelectionCommentTurn(skill, args, msg, turnDeps);
                } else {
                    await runDocumentTurn(skill, args, msg, turnDeps);
                }
                break;
            case 'amendment':
            default:
                if (skill.scope === 'selection-first' && hasSelection) {
                    await runSelectionEditTurn(withArgs(skill.defaultTemplate, args), msg, turnDeps);
                } else {
                    await runDocumentTurn(skill, args, msg, turnDeps);
                }
                break;
        }
    }

    /**
     * Maps a planned task (task-planner.js) to the turn descriptor its
     * pipeline runner expects.
     *
     * @param {{ type: string, instruction: string }} task
     * @param {boolean} hasSelection
     * @returns {{ type: string, instruction?: string, question?: string, scope?: string }}
     * @private
     */
    function turnForTask(task, hasSelection) {
        const instruction = task.instruction;
        // A planner task that is really an empty-paragraph cleanup must not
        // enter the text pipelines — the parser/LLM never see blank
        // paragraphs, so only the deterministic cleanup can serve it,
        // regardless of the planner's own type label (usually "edit").
        if (task.type !== 'qa' && looksLikeCleanupIntent(instruction)) {
            return { type: TURN_TYPE.CLEANUP, instruction };
        }
        switch (task.type) {
            case 'insert':
                // Structural inserts (e.g. a title) belong at document scope —
                // inserting a title into a selection would misplace it.
                return { type: TURN_TYPE.FORMAT, instruction, scope: 'document' };
            case 'format':
                return { type: TURN_TYPE.FORMAT, instruction, scope: hasSelection ? 'selection' : 'document' };
            case 'edit':
                return hasSelection
                    ? { type: TURN_TYPE.SELECTION_EDIT, instruction }
                    : { type: TURN_TYPE.DOC_EDIT, instruction };
            case 'append':
                return { type: TURN_TYPE.DOC_APPEND, instruction };
            case 'table':
                return { type: TURN_TYPE.TABLE, instruction };
            case 'illustration':
                return { type: TURN_TYPE.ILLUSTRATION, instruction };
            case 'qa':
            default:
                return { type: TURN_TYPE.DOC_QA, question: instruction };
        }
    }

    /**
     * Runs a compound turn: the planner decomposes a multi-intent instruction
     * ("增加标题，并深度润色修改") into atomic tasks, then dispatches each to
     * its own pipeline runner — every task stages its own proposal card on
     * this message, in the user-stated order. When planning fails, falls
     * back to single-intent routing of the whole instruction (the
     * pre-planner behavior).
     */
    async function runCompoundTurn(turn, msg, turnDeps, selectionText, turnController) {
        // One shared controller for the whole compound turn: cancel() aborts
        // the in-flight sub-task AND stops the remaining planned tasks.
        const myController = turnController || new AbortController();
        appState.isProcessing = true;
        appState.chatController = myController;
        input.setProcessing(true);
        try {
            msg.setStatus('Planning tasks...');
            const plan = await actions.planDocumentTasks(turnDeps, {
                instruction: turn.instruction,
                hasSelection: !!selectionText,
                signal: myController.signal,
                onToken: (t) => msg.appendModelToken({ id: 'plan' }, 'content', t),
                onReasoning: (t) => msg.appendModelToken({ id: 'plan' }, 'reasoning', t),
            });
            if (!plan.tasks || plan.tasks.length === 0) {
                log('Task planning failed; falling back to single-intent routing.', 'warning');
                const fallback = routeTurn(turn.instruction, {
                    hasSelection: !!selectionText, skills: [], allowCompound: false,
                });
                if (fallback) await dispatchTurn(fallback, msg, turnDeps, selectionText);
                return;
            }
            log(`Executing ${plan.tasks.length} planned task(s): ${plan.tasks.map((t) => t.type).join(' → ')}`, 'info');
            for (let i = 0; i < plan.tasks.length; i++) {
                const task = plan.tasks[i];
                msg.setStatus(`Task ${i + 1}/${plan.tasks.length} [${task.type}]: ${task.instruction}`);
                // Sub-runners toggle isProcessing individually; re-assert the
                // compound turn's busy flag between tasks.
                appState.isProcessing = true;
                input.setProcessing(true);
                await dispatchTurn(turnForTask(task, !!selectionText), msg, turnDeps, selectionText, myController);
                // A cancelled task ends the whole compound turn — the
                // remaining tasks must not start.
                if (myController.signal.aborted) {
                    log(`Compound turn cancelled — skipping ${plan.tasks.length - i - 1} remaining task(s).`, 'warning');
                    break;
                }
            }
            msg.setStatus('');
        } catch (error) {
            if (error.name === 'AbortError') {
                msg.setStatus('Cancelled.');
            } else {
                msg.markError(error.message);
            }
        } finally {
            appState.isProcessing = false;
            if (appState.chatController === myController && !turnController) {
                appState.chatController = null;
            }
            input.setProcessing(false);
        }
    }

    /**
     * Dispatches one routed turn to its pipeline runner. Shared by submit
     * (single turns) and runCompoundTurn (one call per planned task); the
     * optional turnController lets a compound turn share its AbortController
     * with every sub-task, so one cancel stops the whole chain.
     */
    async function dispatchTurn(turn, msg, turnDeps, selectionText, turnController) {
        if (turn.type === TURN_TYPE.SKILL) {
            await runSkillTurn(turn.skill, turn.args, !!selectionText, msg, turnDeps, selectionText, turnController);
        } else if (turn.type === TURN_TYPE.SELECTION_EDIT) {
            await runSelectionEditTurn(turn.instruction, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.DOC_APPEND) {
            await runAppendTurn(turn.instruction, msg, turnDeps, selectionText, turnController);
        } else if (turn.type === TURN_TYPE.ILLUSTRATION) {
            await runIllustrationTurn(turn, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.TABLE) {
            await runTableTurn(turn, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.FORMAT) {
            await runFormatTurn(turn, msg, turnDeps, selectionText, turnController);
        } else if (turn.type === TURN_TYPE.CLEANUP) {
            await runCleanupTurn(msg, turnDeps);
        } else if (turn.type === TURN_TYPE.COMPOUND) {
            await runCompoundTurn(turn, msg, turnDeps, selectionText, turnController);
        } else if (turn.type === TURN_TYPE.DOC_EDIT) {
            // Free-text edit instruction without a selection: run the
            // whole-document amendment pipeline with the user's text as
            // the edit template.
            await runDocumentTurn({
                name: 'Edit', category: 'amendment', scope: 'document',
                defaultTemplate: turn.instruction,
            }, undefined, msg, turnDeps);
        } else {
            await runQaTurn(turn.question, null, msg, turnDeps, selectionText, turnController);
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

        let selectionText = '';
        try {
            selectionText = ((await getSelectionText()) || '').trim();
        } catch (_err) {
            selectionText = '';
        }
        const hasSelection = !!selectionText;

        const turn = routeTurn(trimmed, {
            hasSelection,
            skills: listSkills(appState.promptManager),
        });
        if (!turn) return;

        view.hideWelcome();
        view.addUserMessage(trimmed);
        input.setValue('');

        const msg = view.createAssistantMessage();
        const turnDeps = actionDepsFor(msg);

        try {
            await dispatchTurn(turn, msg, turnDeps, selectionText);
        } catch (error) {
            msg.markError(error.message || String(error));
        } finally {
            // Collapse the per-turn work log and model activity to one-line summaries.
            msg.collapseLog();
            msg.collapseModelOutput();
            // Snapshot the assistant message into the live session array,
            // then notify bootstrap so the session is persisted.
            if (typeof msg.finalizeForHistory === 'function') {
                msg.finalizeForHistory();
            }
            _commitSession();
        }
    }

    /**
     * Cancels the in-flight run (any chat turn or the document pipeline).
     *
     * Both flags/controllers cover all turn types: the document-scope UI
     * flags are released immediately so the user can interact with the
     * input and any pending proposal card without waiting for the
     * in-flight fetches to settle; chat turns (QA, selection edit, append,
     * format, table, illustration, summary, compound planning/sub-tasks)
     * share one AbortController in chatController — aborting it makes the
     * active fetch reject, and each runner's finally releases its own UI
     * state when the rejection lands. Compound turns thread the same
     * controller into every sub-task, so one cancel stops the whole
     * chain instead of just the current task.
     */
    function cancel() {
        let aborted = false;
        if (appState.processDocController) {
            appState.processDocController.abort();
            appState.processDocController = null;
            appState.isProcessingDoc = false;
            input.setProcessing(false);
            aborted = true;
        }
        if (appState.chatController) {
            appState.chatController.abort();
            aborted = true;
        }
        if (aborted) log('Cancelled.', 'warning');
    }

    /**
     * Clears the chat and returns to the welcome state.
     */
    function newChat() {
        if (isBusy()) {
            cancel();
        }
        // Persist the outgoing session BEFORE clearing the live array.
        _commitSession();
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
 * Returns the original text of a chunk result. DocumentChunks carry
 * `paragraphs` (no flat `text` field), so join the paragraph texts; the
 * `.text` fallback covers ad-hoc chunk shapes (e.g. retry payloads).
 *
 * @param {object} result - ChunkResult ({ chunk })
 * @returns {string}
 */
export function chunkOriginalText(result) {
    const c = (result && result.chunk) || {};
    if (Array.isArray(c.paragraphs)) {
        return c.paragraphs.map((p) => ((p && p.text) || '')).join('\n');
    }
    return c.text || '';
}

/**
 * Normalizes text for change detection: CRLF → LF, then trim. Used to decide
 * whether an LLM amendment actually differs from the original chunk text.
 *
 * @param {string} s
 * @returns {string}
 */
function _normalizeText(s) {
    return (s || '').replace(/\r\n/g, '\n').trim();
}

/**
 * Builds citation-pill data for a processed chunk: label and search text are
 * the chunk's first non-empty paragraph (heading or first ~6 words).
 *
 * @param {object} chunk - DocumentChunk ({ id, paragraphs })
 * @returns {{ label: string, searchText: string }}
 */
export function chunkCitation(chunk) {    const firstPara = (chunk.paragraphs || []).map((p) => p.text || '').find((t) => t.trim()) || '';
    const words = firstPara.trim().split(/\s+/).filter(Boolean);
    const label = words.slice(0, 6).join(' ') || chunk.id || 'Section';
    return { label, searchText: firstPara.trim() };
}
