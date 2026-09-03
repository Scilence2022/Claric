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
import { getActiveBackendConfig } from './app-state.js';
import { sendMessages } from '../lib/llm-client.js';
import { connectMcpServer } from '../lib/mcp-client.js';
import { buildLoopTools, createMcpToolExecutor, createResourceClient, RESOURCE_TOOL_SPECS } from '../lib/mcp-tools.js';
import { runToolLoop } from '../lib/tool-loop.js';
import { buildToolLoopSystemPrompt, TOOL_LOOP_LIMITS } from '../lib/tool-registry.js';
import * as agentActions from './agent-actions.js';
import { listSkills, resolveSkill } from './skills.js';
import { createProposalCard as _createProposalCardRaw } from './ui/proposal-card.js';
import { describeFormatOp } from '../lib/format-ops.js';
import { buildAttachmentContext, splitAttachments, attachmentMeta } from '../lib/file-attachments.js';

/** Turn types emitted by routeTurn. */
export const TURN_TYPE = Object.freeze({
    SKILL: 'skill',
    SELECTION_EDIT: 'selection-edit',
    DOC_EDIT: 'doc-edit',
    DOC_APPEND: 'doc-append',
    FORMAT: 'format',
    TABLE: 'table',
    ILLUSTRATION: 'illustration',
    IMAGE_TOOL: 'image-tool',
    TABLE_TOOL: 'table-tool',
    DOCUMENT_IMAGE_TOOL: 'document-image-tool',
    DOCUMENT_TABLE_TOOL: 'document-table-tool',
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
 * True when free text reads as a question (EN + ZH lead, 吗/呢, or a
 * trailing question mark). The 吗/呢 / trailing-? checks catch Chinese
 * yes/no questions that begin with a subject rather than a question word
 * ("你能改变选择的表格的样式吗？") — without them such a question slips into
 * the FORMAT/format-op pipeline and ends as "no changes" instead of an
 * answer.
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeQuestion(text) {
    const trimmed = (text || '').trim();
    if (QUESTION_LEAD_RE.test(trimmed)) return true;
    if (/[吗呢]/.test(trimmed)) return true;
    if (/[？?]$/.test(trimmed)) return true;
    return false;
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
 * match directly; generic image and diagram words (图片/图像/示意图/流程图/
 * image/diagram) require a creation verb nearby, so an edit that merely
 * mentions a picture ("修改图像描述的措辞") still routes to an edit pipeline.
 * These route to the illustration pipeline (image-model or SVG design,
 * then insertion).
 *
 * Diagram nouns deliberately sit in the verb-gated group rather than matching
 * bare: "设计示意图并插入" is a design request, but "把示意图的标题改成..."
 * is an edit, and only the verb tells them apart. The verb list includes bare
 * 画 ("画个流程图") — safe here because a diagram noun must also be present.
 */
const ILLUSTRATION_INTENT_RE = /插图|插画|配图|扉页图|题图|头图|画作|绘制|\bsvg\b|\billustrat|\bartwork\b|(设计|生成|插入|添加|增加|创作|制作|画|做一?[张个幅]).{0,15}(图片|图像|图画|图表|图示|(?:示意|流程|架构|结构|拓扑|原理|思维导|框架|时序|甘特|折线|柱状|饼|网络)图)|(draw|generate|create|add|insert|design|render|make)\b.{0,20}\b(images?|pictures?|drawings?|diagrams?|charts?|figures?|schematics?|flowcharts?)\b/i;

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
 * Image-tool intent markers: the user wants EXISTING images MANAGED —
 * deleted/replaced/resized/relabeled — or SEVERAL illustrations designed at
 * once. Single-image design stays on the dedicated illustration turn (its
 * streaming UX is better); management and multi-image work need the
 * multi-step tool loop. Design verbs alone do not match.
 */
const IMAGE_TOOL_INTENT_RE = /(删除|移除|去掉|替换|更换|重设|缩放|调整|居中|对齐|链接|超链接|比例|标题).{0,8}(图片|图像|插图|配图|插画)|(图片|图像|插图|配图|插画).{0,6}(删除|移除|去掉|替换|更换|太小|太大|大小|尺寸|居中|对齐|链接|超链接|标题|缩放|调整)|(两|三|四|五|几|多)张.{0,4}(插图|配图|插画|图片)|\b(delete|remove|replace|resize|relabel|center|align|link|hyperlink|scal)\w*\b.{0,20}\b(images?|pictures?|illustrations?)|\bset\b.{0,10}\balt\b/i;

/**
 * True when free text asks for image management or multi-image design.
 * Questions stay Q&A; single-image design stays on the illustration turn.
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeImageToolIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return IMAGE_TOOL_INTENT_RE.test(text);
}

/**
 * Document-scope image-management intent markers: the user wants to mutate
 * images across the whole document. Requires an explicit plural-marker
 * ("所有/全部/每个/各/全文/整篇/文档" or English "all/every/each") and a
 * plural-marked noun ("...图片都", "all images"), so a single-image
 * instruction ("把图片居中", "把图片改成…") still routes to the IMAGE_TOOL
 * turn via IMAGE_TOOL_INTENT_RE instead of the planner.
 */
const DOCUMENT_IMAGE_INTENT_RE = /(所有|全部|每个|各|全文|整篇|文档中|文档里).{0,8}(图片|图像|插图|配图|插画)|(图片|图像|插图|配图|插画).{0,4}(全部|都|一起)|\b(all|every|each)\b.{0,12}\b(images?|pictures?|illustrations?)/i;

/**
 * Document-scope table-management intent markers: same plural-marker rule
 * as the image variant — a bare "把表格改成三线表" stays out (it has no
 * plural marker), keeping single-table/selection routing put.
 */
const DOCUMENT_TABLE_INTENT_RE = /(所有|全部|每个|各|全文|整篇|文档中|文档里).{0,8}(表|表格)|(表|表格).{0,4}(全部|都|一起)|\b(all|every|each)\b.{0,12}\btable/i;

/**
 * True when free text (no selection) targets images across the whole document.
 * Questions are excluded so Q&A still routes correctly.
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeDocumentImageIntent(text) {
    const trimmed = (text || '').trim();
    if (looksLikeQuestion(trimmed)) return false;
    return DOCUMENT_IMAGE_INTENT_RE.test(trimmed);
}

/**
 * True when free text (no selection) targets existing tables across the whole
 * document.
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeDocumentTableIntent(text) {
    const trimmed = (text || '').trim();
    if (looksLikeQuestion(trimmed)) return false;
    return DOCUMENT_TABLE_INTENT_RE.test(trimmed);
}

/**
 * Chained-instruction markers: sequence connectors naming MULTIPLE actions
 * ("删除空行，然后重新编号，再加一行合计"). A bare 再 ("再润色一下" — again,
 * one action) does not match; the comma-leading form (，再) does. Chained
 * table instructions route through the table tool loop, which verifies each
 * step against the draft model instead of betting on one-shot JSON.
 */
const CHAIN_RE = /然后|接着|随后|并且|同时|最后|[，、;；]再/;

/**
 * True when the instruction chains multiple actions with sequence
 * connectors.
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeChainedInstruction(text) {
    return CHAIN_RE.test((text || '').trim());
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
 * phrasings plus 行距/列表/编号/项目符号 (spacing and list terms);
 * English terms are word-boundary matched.
 */
const FORMAT_INTENT_RE = /样式|格式|加粗|粗体|斜体|下划线|高亮|标红|字体|字号|颜色|居中|对齐|缩进|行距|间距|列表|编号|项目符号|标题\s*[1-9一二三]?|设为标题|设置为标题|\bbold\b|\bitalic|underline|highlight|font\b|\bcolor\b|\bcenter(ed)?\b|\balign|indent|heading\s*[1-9]|\blists?\b|\bbullets?\b|\bnumbered\b|\bnumbering\b|line\s*spacing|format(ting)?\b/i;

/**
 * Table-look markers inside a format instruction: with a multi-cell table
 * selection these divert the turn to the table tool session, whose style
 * tools (set_table_style / set_borders / set_cell_format / ...) own the
 * table's look. Without a table selection they are ordinary format intent.
 */
const TABLE_STYLE_HINT_RE = /表格|表头|边框|底纹|斑马纹|条纹|列宽|单元格|三线表|隔行|行列|\btable\b|\bborders?\b|\bshading\b|\bcell\b|\bheader row/i;

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
 * Review-intent markers: the user wants the selection ANALYZED or CHECKED,
 * not rewritten ("检查选择的表格的内容"). With a selection this routes to
 * Q&A (the selection becomes focused context) instead of the edit
 * pipelines, which would draft a rewrite. A co-occurring edit verb still
 * wins — "检查并修改这段话" is an edit, not a review.
 */
const REVIEW_INTENT_RE = /\b(check|review|inspect|examine|audit|analy[sz]e|look at)\b|检查|审查|核对|审视|评估|分析|看看/i;

/**
 * True when free text asks for the selection to be analyzed/checked.
 * Questions stay Q&A (already covered by the question-lead check).
 *
 * @param {string} text - Trimmed chat input
 * @returns {boolean}
 */
export function looksLikeReviewIntent(text) {
    if (looksLikeQuestion(text)) return false;
    return REVIEW_INTENT_RE.test(text);
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
 * Filters a tablePatch down to the card-checked items. Checkbox ids run in
 * card order: cells → rowOps → merges → styleOps.
 *
 * @param {object} tablePatch - Proposal's tablePatch
 * @param {Array<number|string>} selectedIds - Checked item ids
 * @returns {object} New patch carrying only the picked ops
 */
function filterTablePatchBySelection(tablePatch, selectedIds) {
    const picked = new Set(selectedIds);
    const cellCount = tablePatch.cells.length;
    const rowOpCount = tablePatch.rowOps.length;
    const mergeCount = (tablePatch.merges || []).length;
    return {
        ...tablePatch,
        cells: tablePatch.cells.filter((_, i) => picked.has(i)),
        rowOps: tablePatch.rowOps.filter((_, j) => picked.has(cellCount + j)),
        merges: (tablePatch.merges || []).filter((_, k) => picked.has(cellCount + rowOpCount + k)),
        styleOps: (tablePatch.styleOps || []).filter((_, s) => picked.has(cellCount + rowOpCount + mergeCount + s)),
    };
}

/**
 * Routes raw chat input to a turn descriptor. Pure function.
 *
 * @param {string} text - Raw chat input
 * @param {object} ctx
 * @param {boolean} ctx.hasSelection - Whether the document has a non-empty
 *   selection (selected text OR selected image(s) OR a multi-cell table
 *   region)
 * @param {boolean} [ctx.hasImageSelection=false] - True when the selection
 *   contains image(s) and NO text: every instruction then enters the image
 *   tool session (object + tools) — questions included, since visual reading
 *   is the read_image tool, not injected bytes
 * @param {boolean} [ctx.hasMultiCellTableRegion=false] - True when the
 *   selection covers multiple cells of a single table (whole table or a
 *   rectangular sub-region): any instruction enters the table tool session
 *   (object + tools). Intra-cell text selections stay on the flat text path.
 * @param {Array<object>} ctx.skills - Available skills (from listSkills)
 * @param {boolean} [ctx.allowCompound=true] - False disables the compound
 *   branch (planner fallback re-routes through single intent)
 * @returns {{ type: string, skill?: object, args?: string, instruction?: string, question?: string, scope?: string } | null}
 *   Null for empty input.
 */
export function routeTurn(text, { hasSelection, hasImageSelection = false, hasMultiCellTableRegion = false, skills, allowCompound = true } = {}) {
    const trimmed = (text || '').trim();
    if (!trimmed) return null;

    const resolved = resolveSkill(trimmed, skills);
    if (resolved) {
        return { type: TURN_TYPE.SKILL, skill: resolved.skill, args: resolved.args };
    }
    // Compound instructions hit several intent families at once; the task
    // planner decomposes them into per-pipeline tasks, executed in order.
    // Document-scope image/table intents ONLY compound when the selection
    // wouldn't already route to IMAGE_TOOL/TABLE_TOOL directly — otherwise
    // a single image selection shouldn't trigger a planner call.
    if (allowCompound) {
        const families = countIntentFamilies(trimmed);
        const docCompound = !hasImageSelection && looksLikeDocumentImageIntent(trimmed) ? 1 : 0;
        const tableDocCompound = !hasMultiCellTableRegion && looksLikeDocumentTableIntent(trimmed) ? 1 : 0;
        if (families + docCompound + tableDocCompound >= 2) {
            // The flag rides along so a planning failure can re-route with the
            // same selection facts (runCompoundTurn's fallback).
            return { type: TURN_TYPE.COMPOUND, instruction: trimmed, hasMultiCellTableRegion };
        }
    }
    // Image-management intent wins over the single-illustration branch:
    // deleting/replacing/resizing existing images or designing several at
    // once needs the multi-step tool loop, not one streaming SVG design.
    if (looksLikeImageToolIntent(trimmed)) {
        return { type: TURN_TYPE.IMAGE_TOOL, instruction: trimmed };
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
    // A format instruction that names the table's look (边框/底纹/表头/...)
    // paired with a multi-cell table selection enters the table tool session
    // instead — its style tools (set_table_style/set_borders/...) own the
    // table. Image-only selections take document scope — format ops target
    // text. Other multi-cell table regions also take document scope — format
    // ops target paragraphs, not table cells.
    if (looksLikeFormatIntent(trimmed)) {
        if (hasMultiCellTableRegion && TABLE_STYLE_HINT_RE.test(trimmed)) {
            return { type: TURN_TYPE.TABLE_TOOL, instruction: trimmed };
        }
        return { type: TURN_TYPE.FORMAT, instruction: trimmed, scope: hasSelection && !hasImageSelection && !hasMultiCellTableRegion ? 'selection' : 'document' };
    }
    // Cleanup intent is document-scope and deterministic: empty paragraphs
    // are invisible to the parser/LLM, so no text pipeline could serve this.
    if (looksLikeCleanupIntent(trimmed)) {
        return { type: TURN_TYPE.CLEANUP, instruction: trimmed };
    }
    if (hasSelection) {
        // Image-only selection: the selection enters as a controllable image
        // OBJECT (snapshot index + metadata + tool list). Questions are
        // answered through read_image inside the session; edits through the
        // op tools — the text pipelines have nothing to operate on.
        if (hasImageSelection) {
            return { type: TURN_TYPE.IMAGE_TOOL, instruction: trimmed };
        }
        // Multi-cell table region: the table enters as a controllable TABLE
        // OBJECT (grid + get_state / set_cell / insert_row / delete_row
        // tools). Mirrors image selection: questions answered inside the
        // session via get_state, edits via the cell/row op tools. Intra-cell
        // text selections skip this branch and stay on the flat text path.
        if (hasMultiCellTableRegion) {
            return { type: TURN_TYPE.TABLE_TOOL, instruction: trimmed };
        }
        // Selection + a question is a question ABOUT the selection (answered
        // in chat with the selection as context), not an edit instruction.
        if (looksLikeQuestion(trimmed)) {
            return { type: TURN_TYPE.DOC_QA, question: trimmed };
        }
        // Selection + review intent is likewise an analysis request, not a
        // rewrite — unless an edit verb co-occurs ("检查并修改").
        if (looksLikeReviewIntent(trimmed) && !looksLikeEditIntent(trimmed)) {
            return { type: TURN_TYPE.DOC_QA, question: trimmed };
        }
        return { type: TURN_TYPE.SELECTION_EDIT, instruction: trimmed };
    }
    // Document-scope image/table intent (no selection, plural-marked):
    // mutates EVERY image/table in the document through the same tool loop,
    // while a selection would have routed to IMAGE_TOOL/TABLE_TOOL above.
    // Runs before the edit branch: "把图片都加上标题" is an image op, not a
    // document amendment.
    if (looksLikeDocumentTableIntent(trimmed)) {
        return { type: TURN_TYPE.DOCUMENT_TABLE_TOOL, instruction: trimmed };
    }
    if (looksLikeDocumentImageIntent(trimmed)) {
        return { type: TURN_TYPE.DOCUMENT_IMAGE_TOOL, instruction: trimmed };
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
        return { type: TURN_TYPE.COMPOUND, instruction: trimmed, hasMultiCellTableRegion };
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
 * @param {function} [deps.getSelectionContent] - async () => ({ text, images })
 *   full selection content at submit time (defaults to word-actions'
 *   readSelectionContent)
 * @param {function} [deps.getSelectionText] - Legacy/test override returning
 *   a plain string (treated as text-only selection)
 * @returns {{ submit: Function, cancel: Function, newChat: Function }}
 */
export function createConversation(deps) {
    const { appState, view, input, log, logWithRetry, updateStatusBar } = deps;
    const actions = deps.actions || { ...defaultActions, ...agentActions };
    // Selection reader: full content ({ text, images }) when available;
    // string-returning overrides (legacy getSelectionText, test mocks)
    // normalize via _normalizeSelection.
    const getSelection = deps.getSelectionContent
        || deps.getSelectionText
        || (typeof actions.readSelectionContent === 'function' ? actions.readSelectionContent : actions.readSelectionSnippet);
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
     * Claims the chat-turn lifecycle: raises the busy flag, registers the
     * turn's AbortController so cancel() reaches it, and locks the input.
     * A compound turn passes its own controller down so one cancel stops the
     * whole chain; every runner must claim through here, or cancel() cannot
     * interrupt it.
     *
     * @param {AbortController} [turnController] - Shared controller from a
     *   compound turn; omitted for a standalone turn
     * @param {'isProcessing'|'isProcessingSummary'} [flag]
     * @returns {AbortController} The controller this turn must use
     * @private
     */
    function _beginChatTurn(turnController, flag = 'isProcessing') {
        const myController = turnController || new AbortController();
        appState[flag] = true;
        appState.chatController = myController;
        input.setProcessing(true);
        return myController;
    }

    /**
     * Releases what _beginChatTurn claimed. The busy flag and the input
     * always clear (an unreleased flag wedges every later turn behind
     * isBusy()); the shared controller is only nulled when this turn owns
     * it — a compound sub-task must leave its parent's controller in place.
     *
     * @param {AbortController} myController
     * @param {AbortController} [turnController]
     * @param {'isProcessing'|'isProcessingSummary'} [flag]
     * @private
     */
    function _endChatTurn(myController, turnController, flag = 'isProcessing') {
        appState[flag] = false;
        if (appState.chatController === myController && !turnController) {
            appState.chatController = null;
        }
        input.setProcessing(false);
    }

    /**
     * Reports a turn failure on the assistant message. Cancellation and the
     * "the model produced no changes" outcome are statuses, not errors.
     *
     * @param {object} msg - Assistant message handle
     * @param {Error} error
     * @private
     */
    function _reportTurnError(msg, error) {
        if (error.name === 'AbortError') {
            msg.setStatus('Cancelled.');
        } else if (error.noChanges) {
            msg.setStatus(error.message);
        } else {
            msg.markError(error.message);
        }
    }

    /**
     * Wraps every proposal card with the shared apply guards: the card's
     * apply AbortController registers with app state (so cancel() reaches
     * it and the busy flags hold) and the input locks for the duration of
     * the apply. Extracted from the document-amendment card — previously
     * only that card wired this, so any other card's apply ran with the
     * busy flags DOWN and could race a new turn's parse → bookmark pass
     * (or a second card's apply). Cards may still override either hook.
     *
     * @param {object} args - createProposalCard args
     * @returns {object} createProposalCard card api
     */
    function makeProposalCard(args) {
        // Tracks the controller this card currently has registered, so a late
        // settle can never release a controller it does not own (e.g. one a
        // follow-up run took over after an in-flight apply was cancelled).
        let ownRegisteredController = null;
        return _createProposalCardRaw({
            isBlocked: () => (isBusy()
                ? 'A run is currently processing — wait for it to finish before applying.'
                : null),
            registerController: (controller) => {
                if (controller === null) {
                    // Apply settled (either the caller settled terminal state
                    // or cancel() already freed the controller) — release the
                    // UI only if the current controller is one THIS card
                    // registered; a foreign run's controller must survive a
                    // late card-apply settle untouched.
                    if (ownRegisteredController
                        && appState.processDocController === ownRegisteredController) {
                        appState.isProcessingDoc = false;
                        appState.processDocController = null;
                        input.setProcessing(false);
                    }
                    ownRegisteredController = null;
                    return;
                }
                // Defensive: never clobber a foreign controller. The card's
                // isBlocked guard should have ensured the busy state was free
                // before we get here; overwriting a run's controller would
                // make the run uncancellable.
                if (appState.processDocController
                    && appState.processDocController !== controller) {
                    return;
                }
                ownRegisteredController = controller;
                appState.processDocController = controller;
                appState.isProcessingDoc = true;
                input.setProcessing(true);
            },
            setApplyBusy: (busy) => {
                input.setProcessing(busy);
            },
            ...args,
        });
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
            if (outcome.failedCount > 0) {
                // Chunks failed (possibly all of them): report the real cause
                // with a retry link instead of masquerading as "the model
                // proposed no changes". The failed chunks' bookmarks stay in
                // place so the retry can target their staged ranges.
                const cancelledNote = outcome.cancelledCount > 0
                    ? ` ${outcome.cancelledCount} section(s) cancelled.`
                    : '';
                const firstError = outcome.results.find((r) => r.status === 'rejected' && r.error)?.error;
                msg.setStatus(
                    `Processing failed on ${outcome.failedCount} of ${outcome.chunks.length} section(s).` +
                    `${cancelledNote} No changes were applied.` +
                    (firstError ? ` Last error: ${firstError}` : '')
                );
                if (typeof outcome.retryFailed === 'function' && typeof logWithRetry === 'function') {
                    logWithRetry(
                        `${outcome.failedCount} section(s) failed to process. Click to retry.`,
                        'warning',
                        // The retry fires from the activity log long after this
                        // turn returned, so it owns its own input lock — hand it
                        // the setter, since word-actions cannot reach the input.
                        () => outcome.retryFailed({ setBusy: (busy) => input.setProcessing(busy) })
                    );
                }
                return;
            }
            await outcome.discard();
            msg.setStatus(outcome.cancelledCount > 0
                ? 'Cancelled — no changes were applied.'
                : 'The model proposed no changes.');
            return;
        }

        const beforeChars = amendedChunks.reduce((s, r) => s + chunkOriginalText(r).length, 0);
        const afterChars = amendedChunks.reduce((s, r) => s + (r.amendment || '').length, 0);
        msg.setStatus('');

        const card = makeProposalCard({
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
            onApply: async (selectedChunkIds, applyCtx = {}) => {
                appState.isProcessingDoc = true;
                input.setProcessing(true);
                try {
                    const applicationResult = await outcome.apply(selectedChunkIds, applyCtx);
                    if (applicationResult.interrupted) {
                        // Stopped mid-apply (Stop button): the remaining chunks
                        // keep their bookmarks; re-enable "Continue applying".
                        const appliedCount = applicationResult.appliedChunkIds.length;
                        const total = selectedChunkIds ? selectedChunkIds.length : outcome.chunks.length;
                        card.setPaused(
                            `Paused — ${appliedCount} of ${total} selected change(s) applied. Click "Continue applying" to resume.`
                        );
                        msg.setStatus('');
                        return;
                    }
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
                    // Busy flags are NOT reset here: the card's apply handler
                    // releases them via registerController(null), which checks
                    // that we still own the shared controller — an
                    // unconditional reset used to clobber a newer run's state
                    // when cancel() had already freed the UI.
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
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Drafting edit...');
            const commentInstructions = getCommentInstructions();

            let proposal;
            // Chained instructions on (possibly) a table selection take the
            // tool loop: each step is validated against the draft model
            // instead of betting on one-shot JSON. Non-table selections
            // resolve to null and fall through to the single-shot path.
            if (looksLikeChainedInstruction(promptTemplate)) {
                msg.setStatus('Working through the steps (tool loop)...');
                proposal = await actions.prepareTableToolEdit(turnDeps, {
                    instruction: promptTemplate,
                    signal: myController.signal,
                    onStep: (s) => msg.appendModelToken({ id: 'tools' }, 'content',
                        s.text ? `${s.text}\n` : ''),
                });
            }
            if (!proposal) {
                try {
                    proposal = await actions.prepareSelectionAmendment(turnDeps, {
                        promptTemplate,
                        commentInstructions,
                        signal: myController.signal,
                        onToken: (t) => msg.appendModelToken({ id: 'selection' }, 'content', t),
                        onReasoning: (t) => msg.appendModelToken({ id: 'selection' }, 'reasoning', t),
                    });
                } catch (err) {
                    // A table selection whose one-shot JSON patch didn't
                    // parse gets one structured retry via the tool loop —
                    // step-wise validation recovers what one-shot couldn't.
                    if (/no JSON object/i.test(err.message)) {
                        log(`Single-shot table patch failed (${err.message}); retrying via the tool loop.`, 'warning');
                        proposal = await actions.prepareTableToolEdit(turnDeps, {
                            instruction: promptTemplate,
                            signal: myController.signal,
                            onStep: (s) => msg.appendModelToken({ id: 'tools' }, 'content',
                                s.text ? `${s.text}\n` : ''),
                        });
                        if (!proposal) throw err;
                    } else {
                        throw err;
                    }
                }
            }
            // Read-only tool-loop outcome (the model inspected the table and
            // answered instead of editing): the summary IS the chat answer.
            // Staging a card here would have nothing to write, yet report
            // "Applied as tracked changes" on apply.
            if (proposal.noOps) {
                msg.setStatus('');
                msg.setText(proposal.answer || '(no changes)');
                return;
            }
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

            const card = makeProposalCard({
                title,
                beforeChars,
                afterChars,
                comment: proposal.commentText,
                items,
                onLocate: (text) => actions.revealTextSnippet(turnDeps, text),
                onApply: async (selectedIds) => {
                    try {
                        // Table cards list one checkbox per cell patch, then
                        // per row op, merge, and style op — honor unchecking.
                        const toApply = (isTable && selectedIds)
                            ? { ...proposal, tablePatch: filterTablePatchBySelection(proposal.tablePatch, selectedIds) }
                            : proposal;
                        const applied = await actions.applySelectionAmendment(turnDeps, toApply);
                        if (applied && applied.skipped) {
                            // Staleness guard refused the write — surface it
                            // as a card warning, not a success.
                            card.markWarning(applied.reason);
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
                countsText: `${beforeChars} → ${afterChars} chars`,
                items: items || [],
            });
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs an append turn: the LLM drafts new content against the document
     * context, staged in a proposal card. Apply inserts it at the document
     * end as tracked changes; nothing is written before that.
     */
    async function runAppendTurn(instruction, msg, turnDeps, selectionText, turnController) {
        const myController = _beginChatTurn(turnController);
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
            const card = makeProposalCard({
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
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
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
        const myController = _beginChatTurn(turnController);
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
            const card = makeProposalCard({
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
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
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
        const myController = _beginChatTurn(turnController);
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
            const card = makeProposalCard({
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
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs an illustration turn: the LLM designs one self-contained SVG from
     * the document's subject and mood, staged in a proposal card with an
     * image preview. Apply inserts the returned raster bytes into the
     * document via Word.js (tracked when track-changes is on); SVG fallback
     * proposals are rasterized to PNG at apply time.
     */
    async function runIllustrationTurn(turn, msg, turnDeps, turnController) {
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Designing illustration...');
            const proposal = await actions.prepareIllustrationProposal(turnDeps, {
                instruction: turn.instruction,
                signal: myController.signal,
                onToken: (t) => msg.appendModelToken({ id: 'illustration' }, 'content', t),
                onReasoning: (t) => msg.appendModelToken({ id: 'illustration' }, 'reasoning', t),
            });
            if (!proposal.svg && !proposal.imageBase64) {
                msg.setStatus('The model produced no usable illustration.');
                return;
            }
            msg.setStatus('');
            const positionLabel = proposal.positionLabel || `document ${proposal.position}`;
            const countsText = `${proposal.sizeLabel || 'illustration'} at ${positionLabel}`;
            // The image engine returns raster bytes rendered through a
            // MIME-aware data URL; the SVG engine returns sanitized markup
            // rendered inline. Only one is ever set — see
            // prepareIllustrationProposal.
            const preview = proposal.imageBase64
                ? { previewSrc: proposal.previewSrc || `data:image/png;base64,${proposal.imageBase64}` }
                : { previewSvg: proposal.svg };
            const card = makeProposalCard({
                title: 'Proposed illustration',
                countsText,
                ...preview,
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
                countsText,
                ...preview,
                items: [],
            });
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs an image-management tool-loop turn: the model drives list/design/
     * replace/delete/resize/alt-text tools against a draft snapshot of the
     * document's images, and the recorded ops stage in a proposal card with
     * one selectable change per op. Apply runs only the checked ops.
     */
    async function runImageToolTurn(turn, msg, turnDeps, selectionImages, turnController) {
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Working through the image task (tool loop)...');
            const proposal = await actions.prepareImageToolEdit(turnDeps, {
                instruction: turn.instruction,
                selectionImages,
                signal: myController.signal,
                onStep: (s) => msg.appendModelToken({ id: 'tools' }, 'content',
                    s.text ? `${s.text}\n` : ''),
            });
            _stageImageOpsProposal(proposal, msg, turnDeps, {
                title: 'Proposed image changes',
                rejectLog: 'Image proposal rejected by user.',
                emptyStatus: 'The image snapshot is no longer available — nothing to operate on.',
            });
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Stages an image tool-loop result: renders a read-only outcome as the
     * chat answer, else builds the one-checkbox-per-op proposal card whose
     * apply runs only the checked ops. Shared by the selection-scope and
     * document-scope image turns — they differ only in the labels.
     *
     * @param {object|null} proposal - prepareImageToolEdit result
     * @param {object} msg - Assistant message handle
     * @param {object} turnDeps
     * @param {{title: string, rejectLog: string, emptyStatus: string}} labels
     * @private
     */
    function _stageImageOpsProposal(proposal, msg, turnDeps, labels) {
        if (!proposal) {
            msg.setStatus(labels.emptyStatus);
            return;
        }
        // Read-only outcome (e.g. the model inspected the image with
        // read_image and answered): the finish summary is the chat answer —
        // no proposal card, nothing to apply.
        if (proposal.noOps) {
            msg.setStatus('');
            msg.setText(proposal.answer || '(no image changes)');
            return;
        }
        msg.setStatus('');
        const { title } = labels;
        const countsText = `${proposal.ops.length} image operation(s)`;
        const svgOps = proposal.items.filter((item) => item.svg);
        // Already sanitized by the image tool loop (design/replace_illustration
        // ops run sanitizeSvg + ensureSvgDimensions). Inline render avoids the
        // SVG data-URL decode that fails on some hosts.
        const previewSvg = svgOps.length === 1 ? svgOps[0].svg : undefined;
        const cardItems = proposal.items.map(({ id, label, before, after }) => ({
            id, label, before, after,
        }));
        const card = makeProposalCard({
            title,
            countsText,
            previewSvg,
            comment: null,
            items: cardItems,
            onApply: async (selectedIds) => {
                try {
                    // One checkbox per op — honor the user's unchecking.
                    const picked = new Set(selectedIds);
                    const ops = proposal.ops.filter((_, i) => picked.has(i + 1));
                    if (ops.length === 0) {
                        card.markWarning('No operations selected.');
                        return;
                    }
                    const result = await actions.applyImageOps(turnDeps, { ...proposal, ops });
                    if (result && result.warnings && result.warnings.length > 0) {
                        card.markWarning(`Applied with warning: ${result.warnings[0]}`);
                    } else {
                        card.markApplied();
                    }
                } catch (error) {
                    log(`Apply failed: ${error.message}`, 'error');
                    card.markError(error.message);
                }
            },
            onReject: () => {
                log(labels.rejectLog, 'info');
            },
        });
        msg.attachProposal(card, {
            title,
            state: 'pending',
            countsText,
            previewSvg,
            items: cardItems,
        });
    }

    /**
     * Stages a table tool-loop result: renders a read-only outcome as the
     * chat answer, else builds the per-cell/row/merge/style checkbox card
     * whose apply filters the patch to the checked items. Shared by the
     * selection-scope and document-scope table turns.
     *
     * @param {object|null} proposal - prepareTableToolEdit result
     * @param {object} msg - Assistant message handle
     * @param {object} turnDeps
     * @param {{title: string, rejectLog: string, emptyStatus: string}} labels
     * @private
     */
    function _stageTablePatchProposal(proposal, msg, turnDeps, labels) {
        if (!proposal) {
            msg.setStatus(labels.emptyStatus);
            return;
        }
        if (proposal.noOps) {
            msg.setStatus('');
            msg.setText(proposal.answer || '(no table changes)');
            return;
        }
        msg.setStatus('');
        const { title } = labels;
        const beforeChars = proposal.tableItems.reduce((sum, item) => sum + (item.before || '').length, 0);
        const afterChars = proposal.tableItems.reduce((sum, item) => sum + (item.after || '').length, 0);
        const items = proposal.tableItems.map((item, index) => ({
            id: index,
            label: item.label,
            before: item.before,
            after: item.after,
            searchText: item.searchText,
        }));
        const card = makeProposalCard({
            title,
            beforeChars,
            afterChars,
            comment: null,
            items,
            onLocate: (text) => actions.revealTextSnippet(turnDeps, text),
            onApply: async (selectedIds) => {
                try {
                    const toApply = selectedIds
                        ? { ...proposal, tablePatch: filterTablePatchBySelection(proposal.tablePatch, selectedIds) }
                        : proposal;
                    const applied = await actions.applySelectionAmendment(turnDeps, toApply);
                    if (applied && applied.skipped) {
                        // Staleness guard refused the write — a warning, not
                        // a success.
                        card.markWarning(applied.reason);
                    } else {
                        card.markApplied();
                    }
                } catch (error) {
                    log(`Apply failed: ${error.message}`, 'error');
                    card.markError(error.message);
                }
            },
            onReject: () => {
                log(labels.rejectLog, 'info');
            },
        });
        msg.attachProposal(card, {
            title,
            state: 'pending',
            countsText: `${beforeChars} → ${afterChars} chars`,
            items,
        });
    }

    /**
     * Runs a table tool turn for a multi-cell table selection: every
     * instruction enters the selection as a controllable TABLE OBJECT
     * (grid + get_state / set_cell / insert_row / delete_row tools). The
     * loop's translated patch reuses the existing tablePatch proposal
     * shape, so the same per-cell checkbox card + applySelectionAmendment
     * (table branch) serve unchanged. Mirrors runImageToolTurn: a
     * read-only outcome (loop finishes with summary, no ops) renders the
     * summary as chat text — no card.
     */
    async function runTableToolTurn(turn, msg, turnDeps, turnController) {
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Working through the table task (tool loop)...');
            const proposal = await actions.prepareTableToolEdit(turnDeps, {
                instruction: turn.instruction,
                signal: myController.signal,
                onStep: (s) => msg.appendModelToken({ id: 'tools' }, 'content',
                    s.text ? `${s.text}\n` : ''),
            });
            _stageTablePatchProposal(proposal, msg, turnDeps, {
                title: 'Proposed table edit',
                rejectLog: 'Table proposal rejected.',
                emptyStatus: 'The selection is no longer a multi-cell table region — nothing to operate on.',
            });
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs a cleanup turn: deterministically scans the document for redundant
     * empty paragraphs (no LLM — blank paragraphs are invisible to the parser
     * and excluded from text-pipeline alignment, so a Word.js scan is the only
     * way to serve this) and stages a proposal card. Apply deletes them as
     * tracked changes.
     */
    async function runCleanupTurn(msg, turnDeps, turnController) {
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Scanning for empty paragraphs...');
            const proposal = await actions.prepareEmptyParagraphCleanup(turnDeps);
            if (!proposal.emptyCount) {
                msg.setStatus('No empty paragraphs found.');
                return;
            }
            msg.setStatus('');
            const card = makeProposalCard({
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
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs a selection-scope comment turn (fire-and-forget comment pipeline).
     *
     * Claims the turn lifecycle like every other runner so cancel() can reach
     * the in-flight LLM request (see _beginChatTurn). The claim is released as
     * soon as the request is QUEUED, not when the comment lands: the pipeline
     * is deliberately fire-and-forget so several comments can be pending at
     * once, and holding the busy flag until insertion would serialize them.
     * The signal keeps working after release — an abort still suppresses the
     * insert (comment-request.js).
     */
    async function runSelectionCommentTurn(skill, args, msg, turnDeps, turnController) {
        if (!appState.supportsComments) {
            msg.markError('Comment features require Word API 1.4.');
            return;
        }
        const myController = _beginChatTurn(turnController);
        try {
            await actions.fireSelectionComment(turnDeps, {
                promptTemplate: withArgs(skill.defaultTemplate, args),
                signal: myController.signal,
            });
            msg.setStatus('Comment request fired — it will appear in the document shortly.');
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs the summary pipeline (result opens as a new document).
     */
    async function runSummaryTurn(skill, args, msg, turnDeps, turnController) {
        const myController = _beginChatTurn(turnController, 'isProcessingSummary');
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
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController, 'isProcessingSummary');
        }
    }

    /**
     * Runs a chat Q&A turn with streaming.
     * selectionText (when non-empty) is added to the prompt as a focused
     * excerpt alongside the full document context; selectionImages metadata
     * rides along as object references for mixed text+image selections.
     * questionImages (chat file-upload attachments) carry dataUrls and are
     * sent to the model as image_url parts, with a text-only fallback when
     * the backend rejects image inputs.
     */
    async function runQaTurn(question, skillTemplate, msg, turnDeps, selectionText, selectionImages, turnController, questionImages) {
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Reading the document...');
            const answer = await actions.answerQuestion(turnDeps, {
                question,
                skillTemplate,
                selectionText,
                selectionImages,
                questionImages,
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
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Dispatches a skill turn by category and resolved scope.
     */
    /**
     * Reserved /mcp turn: a ReAct tool loop over the configured MCP
     * servers. Read-only contract — tool results render as the chat
     * answer; nothing writes to the Word document from this path (MCP
     * tools act on their own external systems).
     *
     * @param {string} args - Instruction after /mcp
     * @param {object} msg - Assistant message view
     * @param {string} selectionText - Current selection (focus context)
     * @param {AbortController} turnController - Shared turn controller
     * @private
     */
    async function runMcpToolsTurn(args, msg, selectionText, turnController) {
        const instruction = (args || '').trim() ||
            'List the tools available to you and summarize what you can do with them.';
        const servers = (appState.config.mcpServers || [])
            .filter((s) => s && s.url && s.enabled !== false);
        if (servers.length === 0) {
            msg.appendText('No MCP servers are configured. Add one in Settings → MCP Servers, then run /mcp again.');
            return;
        }
        const budget = Number(appState.config.mcpStepBudget) > 0
            ? Math.round(Number(appState.config.mcpStepBudget))
            : TOOL_LOOP_LIMITS.MAX_STEPS_DEFAULT;

        log(`Connecting to ${servers.length} MCP server(s)...`, 'info');
        const connected = [];
        for (const server of servers) {
            const label = server.name || server.url;
            try {
                const client = await connectMcpServer({ url: server.url, token: server.token, log });
                const mcpTools = await client.listTools();
                connected.push({ name: label, client, mcpTools });
                log(`MCP "${label}": ${mcpTools.length} tool(s) available.`, 'info');
            } catch (err) {
                log(`MCP server "${label}" failed: ${err.message}`, 'warning');
            }
        }
        if (connected.length === 0) {
            msg.appendText('No MCP server could be reached — see the activity log for details.');
            return;
        }

        // The resource pseudo-server exposes mcp_list_resources /
        // mcp_read_resource so the model can pull reference material itself.
        connected.push({ name: 'resources', client: createResourceClient(connected), mcpTools: RESOURCE_TOOL_SPECS });
        const { loopTools, mapping } = buildLoopTools(connected);
        if (loopTools.length <= RESOURCE_TOOL_SPECS.length) {
            msg.appendText('The configured MCP servers expose no tools.');
            return;
        }

        const myController = _beginChatTurn(turnController);
        try {
            const config = getActiveBackendConfig(appState);
            const taskPrompt = instruction + (selectionText
                ? `\n\n--- SELECTED TEXT (focus) ---\n${selectionText}`
                : '');
            msg.setStatus('Working with MCP tools...');

            const result = await runToolLoop({
                systemPrompt: buildToolLoopSystemPrompt(loopTools, { maxSteps: budget }),
                taskPrompt,
                tools: loopTools,
                execute: createMcpToolExecutor(mapping),
                send: (messages) => sendMessages(config, messages, log, myController.signal, 120000),
                maxSteps: budget,
                signal: myController.signal,
                onStep: (step) => {
                    if (step.call) {
                        msg.appendLogLine(`MCP step ${step.step}: ${step.call.tool}${step.ok === false ? ' (failed)' : ''}`);
                    }
                },
            });
            msg.setStatus('');
            msg.setText(result.summary || 'Done — the tool loop finished without a summary.');
        } catch (error) {
            if (error.name !== 'AbortError') {
                log(`MCP tools turn failed: ${error.message}`, 'error');
            }
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    async function runSkillTurn(skill, args, hasSelection, msg, turnDeps, selectionText, selectionImages, turnController) {
        switch (skill.category) {
            case 'chat':
            case 'context':
                // Custom context prompts act as chat personas.
                await runQaTurn(args || skill.description, skill.defaultTemplate, msg, turnDeps, selectionText, selectionImages, turnController);
                break;
            case 'summary':
                await runSummaryTurn(skill, args, msg, turnDeps, turnController);
                break;
            case 'tools':
                // Reserved /mcp skill: a ReAct tool loop over the configured
                // MCP servers. Read-only contract — results render as the
                // chat answer, never as direct document writes.
                await runMcpToolsTurn(args, msg, selectionText, turnController);
                break;
            case 'comment':
                if (skill.scope === 'selection-first' && hasSelection) {
                    await runSelectionCommentTurn(skill, args, msg, turnDeps, turnController);
                } else {
                    await runDocumentTurn(skill, args, msg, turnDeps);
                }
                break;
            case 'amendment':
            default:
                if (skill.scope === 'selection-first' && hasSelection) {
                    await runSelectionEditTurn(withArgs(skill.defaultTemplate, args), msg, turnDeps, turnController);
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
            case 'image_management':
                // Whole-document image tool session — separate card; user
                // reviews & selectively applies alongside any text edits.
                return { type: TURN_TYPE.DOCUMENT_IMAGE_TOOL, instruction };
            case 'table_management':
                // First table in the document (v1 limitation; follow-up
                // turns can target additional tables).
                return { type: TURN_TYPE.DOCUMENT_TABLE_TOOL, instruction };
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
    async function runCompoundTurn(turn, msg, turnDeps, selectionText, selectionImages, turnController) {
        // One shared controller for the whole compound turn: cancel() aborts
        // the in-flight sub-task AND stops the remaining planned tasks.
        const myController = _beginChatTurn(turnController);
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
                // Same three selection facts submit() routed on — dropping
                // the table-region flag would re-route a multi-cell table
                // selection as if nothing were selected.
                const hasTableRegion = !!turn.hasMultiCellTableRegion;
                const fallback = routeTurn(turn.instruction, {
                    hasSelection: !!selectionText || selectionImages.length > 0 || hasTableRegion,
                    hasImageSelection: !selectionText && !hasTableRegion && selectionImages.length > 0,
                    hasMultiCellTableRegion: hasTableRegion,
                    skills: [], allowCompound: false,
                });
                if (fallback) await dispatchTurn(fallback, msg, turnDeps, selectionText, selectionImages, myController);
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
                await dispatchTurn(turnForTask(task, !!selectionText), msg, turnDeps, selectionText, selectionImages, myController);
                // A cancelled task ends the whole compound turn — the
                // remaining tasks must not start.
                if (myController.signal.aborted) {
                    log(`Compound turn cancelled — skipping ${plan.tasks.length - i - 1} remaining task(s).`, 'warning');
                    break;
                }
            }
            msg.setStatus('');
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs a document-scope image management turn: every inline picture in
     * the document is snapshotted into an operable object with the full
     * `IMAGE_TOOL_SPECS` tool list. Mirrors runImageToolTurn (selection
     * variant) but skips the selection-image gate — the user is talking
     * about the document as a whole.
     *
     * @private
     */
    async function runDocumentImageToolTurn(turn, msg, turnDeps, turnController) {
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Working through document images (tool loop)...');
            const proposal = await actions.prepareImageToolEdit(turnDeps, {
                instruction: turn.instruction,
                signal: myController.signal,
                onStep: (s) => msg.appendModelToken({ id: 'tools' }, 'content',
                    s.text ? `${s.text}\n` : ''),
            });
            const pictures = proposal
                ? `${proposal.snapshotCount} picture${proposal.snapshotCount === 1 ? '' : 's'}`
                : '';
            _stageImageOpsProposal(proposal, msg, turnDeps, {
                title: `Document image changes (${pictures})`,
                rejectLog: 'Document image proposal rejected by user.',
                emptyStatus: 'Document image snapshot failed.',
            });
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Runs a document-scope table management turn: EVERY table in the
     * document becomes an operable object (indexed by `tableIndex`) with the
     * full `TABLE_TOOL_SPECS` tool list. The loop may cross tables in one
     * session — e.g. "把所有表格改成三线表,并把第二个表的表头加底纹".
     *
     * @private
     */
    async function runDocumentTableToolTurn(turn, msg, turnDeps, turnController) {
        const myController = _beginChatTurn(turnController);
        try {
            msg.setStatus('Working through document tables (tool loop)...');
            const regions = await actions.readDocumentTableRegions(turnDeps);
            if (!regions || regions.length === 0) {
                msg.setStatus('Document has no tables to manage — nothing to operate on.');
                return;
            }
            const proposal = await actions.prepareTableToolEdit(turnDeps, {
                instruction: turn.instruction,
                regions,
                signal: myController.signal,
                onStep: (s) => msg.appendModelToken({ id: 'tools' }, 'content',
                    s.text ? `${s.text}\n` : ''),
            });
            _stageTablePatchProposal(proposal, msg, turnDeps, {
                title: 'Document table changes',
                rejectLog: 'Document table proposal rejected.',
                emptyStatus: 'Document has no tables to manage — nothing to operate on.',
            });
        } catch (error) {
            _reportTurnError(msg, error);
        } finally {
            _endChatTurn(myController, turnController);
        }
    }

    /**
     * Dispatches one routed turn to its pipeline runner. Shared by submit
     * (single turns) and runCompoundTurn (one call per planned task); the
     * optional turnController lets a compound turn share its AbortController
     * with every sub-task, so one cancel stops the whole chain.
     */
    async function dispatchTurn(turn, msg, turnDeps, selectionText, selectionImages, turnController) {
        if (turn.type === TURN_TYPE.SKILL) {
            await runSkillTurn(turn.skill, turn.args, !!selectionText, msg, turnDeps, selectionText, selectionImages, turnController);
        } else if (turn.type === TURN_TYPE.SELECTION_EDIT) {
            await runSelectionEditTurn(turn.instruction, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.DOC_APPEND) {
            await runAppendTurn(turn.instruction, msg, turnDeps, selectionText, turnController);
        } else if (turn.type === TURN_TYPE.ILLUSTRATION) {
            await runIllustrationTurn(turn, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.IMAGE_TOOL) {
            await runImageToolTurn(turn, msg, turnDeps, selectionImages, turnController);
        } else if (turn.type === TURN_TYPE.TABLE) {
            await runTableTurn(turn, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.TABLE_TOOL) {
            await runTableToolTurn(turn, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.DOCUMENT_IMAGE_TOOL) {
            await runDocumentImageToolTurn(turn, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.DOCUMENT_TABLE_TOOL) {
            await runDocumentTableToolTurn(turn, msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.FORMAT) {
            await runFormatTurn(turn, msg, turnDeps, selectionText, turnController);
        } else if (turn.type === TURN_TYPE.CLEANUP) {
            await runCleanupTurn(msg, turnDeps, turnController);
        } else if (turn.type === TURN_TYPE.COMPOUND) {
            await runCompoundTurn(turn, msg, turnDeps, selectionText, selectionImages, turnController);
        } else if (turn.type === TURN_TYPE.DOC_EDIT) {
            // Free-text edit instruction without a selection: run the
            // whole-document amendment pipeline with the user's text as
            // the edit template.
            await runDocumentTurn({
                name: 'Edit', category: 'amendment', scope: 'document',
                defaultTemplate: turn.instruction,
            }, undefined, msg, turnDeps);
        } else {
            await runQaTurn(turn.question, null, msg, turnDeps, selectionText, selectionImages, turnController, turn.questionImages);
        }
    }

    /**
     * Submits raw chat input as a new turn.
     *
     * @param {string} text
     * @param {Array<{name: string, kind: string, size: number, text?: string, dataUrl?: string}>} [attachments] -
     *   Parsed files from the input bar. Extracted text (txt/md/docx/pdf) is
     *   appended to the routed question/instruction as labeled sections;
     *   images ride QA turns as image_url parts (turn.questionImages).
     */
    async function submit(text, attachments) {
        const list = Array.isArray(attachments) ? attachments.filter(Boolean) : [];
        const trimmed = (text || '').trim();
        if (!trimmed && list.length === 0) return;
        // Question lead so attachment-only submissions route to DOC_QA, not
        // the compound planner.
        const effective = trimmed || 'What do the attached file(s) say?';

        if (isBusy()) {
            log('Already processing. Cancel the current run first.', 'warning');
            return;
        }

        let selectionText = '';
        // Metadata only — base64 payloads never leave the selection readers
        // (preview thumbnails come from watchSelection's own read).
        let selectionImages = [];
        // Multi-cell table region: matched in readSelectionContent's same
        // Word.run — boolean flag for routing, full coords available via
        // the full read inside prepareTableToolEdit.
        let hasMultiCellTableRegion = false;
        try {
            const raw = await getSelection();
            const sel = _normalizeSelection(raw);
            const extras = (raw && typeof raw === 'object') ? raw : {};
            selectionText = (sel.text || '').trim();
            selectionImages = sel.images;
            hasMultiCellTableRegion = !!extras.hasMultiCellTableRegion;
        } catch (_err) {
            selectionText = '';
            selectionImages = [];
            hasMultiCellTableRegion = false;
        }
        const hasSelection = !!selectionText || selectionImages.length > 0 || hasMultiCellTableRegion;
        const hasImageSelection = !selectionText && !hasMultiCellTableRegion && selectionImages.length > 0;

        const turn = routeTurn(effective, {
            hasSelection,
            hasImageSelection,
            hasMultiCellTableRegion,
            skills: listSkills(appState.promptManager),
        });
        if (!turn) return;

        // Attachments: extracted text joins the routed prompt as labeled
        // sections; images become image_url parts on QA turns (other
        // pipelines are text-only — the context block lists them by name).
        if (list.length > 0) {
            const context = buildAttachmentContext(list);
            if (context) {
                if (typeof turn.question === 'string') turn.question += context;
                else if (typeof turn.instruction === 'string') turn.instruction += context;
            }
            turn.questionImages = splitAttachments(list).imageAttachments;
        }

        view.hideWelcome();
        view.addUserMessage(effective, attachmentMeta(list));
        input.setValue('');

        const msg = view.createAssistantMessage();
        const turnDeps = actionDepsFor(msg);

        try {
            await dispatchTurn(turn, msg, turnDeps, selectionText, selectionImages);
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
        if (typeof input.clearAttachments === 'function') input.clearAttachments();
        input.focus();
    }

    return { submit, cancel, newChat };
}

/**
 * Normalizes a selection-reader result into { text, images }. Legacy
 * string readers (getSelectionText overrides, test mocks) map to text-only;
 * object results carry image metadata ({ width, height, altText }) with any
 * base64 payload stripped.
 *
 * @param {string|{text?: string, images?: Array<object>}} result
 * @returns {{ text: string, images: Array<{width: number, height: number, altText: string}> }}
 */
function _normalizeSelection(result) {
    if (typeof result === 'string') return { text: result, images: [] };
    if (result && Array.isArray(result.images)) {
        return {
            text: result.text || '',
            images: result.images.map(({ width, height, altText }) => ({ width, height, altText: altText || '' })),
        };
    }
    return { text: '', images: [] };
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
export function chunkCitation(chunk) {
    const firstPara = (chunk.paragraphs || []).map((p) => p.text || '').find((t) => t.trim()) || '';
    const words = firstPara.trim().split(/\s+/).filter(Boolean);
    const label = words.slice(0, 6).join(' ') || chunk.id || 'Section';
    return { label, searchText: firstPara.trim() };
}
