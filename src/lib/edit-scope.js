/**
 * Edit Scope Module
 *
 * Scope control for the chunked whole-document amendment pipeline, which
 * sends every chunk with the user's instruction. "Rewrite only X, change
 * nothing else" would otherwise still change whatever else the model touches
 * in other chunks. {@link parseEditScope} recognises an instruction that names
 * a specific target, {@link locateScopeTargets} finds it in the document, and
 * {@link confineAmendment} / {@link confineChunkResults} project replies onto
 * it so no other paragraph can change, whatever the model returned.
 *
 * Loaded lazily (task-module-loader.js loadEditScope) so it stays out of the
 * task pane's start-up bundle. Pure JavaScript -- no Word API dependency.
 *
 * @module edit-scope
 */

import { alignParagraphs } from './paragraph-alignment.js';
import { intentText } from './intent-text.js';
import { amendedLines, summarizeAmendment } from './amendment-guard.js';

/**
 * @typedef {Object} EditScope
 * @property {boolean} restricted - The instruction limits the edit to a named part of the document
 * @property {Array<{name: string, noun: string}>} targets - Named targets ("Abstract", "paragraph")
 * @property {boolean} positional - Names a position ("paragraph 3", "第三段") that needs a human-style count
 */

const NOUN_EN = '(?:paragraph|section|chapter|subsection|part|passage|portion|heading)';
// 1-4 words naming the target ("Abstract", "literature review"), optionally quoted.
// A name never spans a preposition or article: "grammar in the Methods section"
// names Methods, not "grammar in the Methods".
const NAME_WORD_EN = '(?!(?:in|within|inside|under|the|to|for|with|on|at|by|so|that|which|is|are|be)\\b)[A-Za-z0-9][\\w&\'’\\-/]*';
const NAME_EN = `["“‘']?(${NAME_WORD_EN}(?:\\s+${NAME_WORD_EN}){0,3})["”’']?`;
const EDIT_VERB_EN = '(?:rewrite|revise|polish|edit|shorten|expand|condense|improve|fix|proofread|translate|reword|rephrase|tighten|simplify|modify|change|update|correct|refine|summari[sz]e|rework|redraft|copy-?edit|amend|trim|lengthen|enrich|strengthen|delete|remove)';
const SECTION_NAMES_EN = '(?:abstract|executive summary|summary|introduction|background|methods?|methodology|results|discussion|conclusions?|recommendations?|acknowledge?ments?|overview|literature review|limitations|appendix|references)';

/** Words that name no particular part: "the other section", "the following paragraph". */
const GENERIC_NAMES = new Set([
  'other', 'another', 'any', 'all', 'each', 'every', 'whole', 'entire', 'rest', 'remaining', 'same', 'this', 'that',
  'these', 'those', 'following', 'previous', 'next', 'selected', 'given', 'current', 'above', 'below', 'first', 'second',
  'third', 'fourth', 'fifth', 'last', 'final', 'document', 'article', 'text', 'paper', 'one', 'new', 'a', 'an',
]);

const TARGET_PATTERNS_EN = [
  // "rewrite only the Abstract paragraph", "polish the Methods section"
  { re: new RegExp(`\\b${EDIT_VERB_EN}\\s+(?:only\\s+|just\\s+)?(?:the\\s+)?${NAME_EN}\\s+(${NOUN_EN})\\b`, 'gi'), needsExclusive: false },
  // "only the Abstract paragraph"
  { re: new RegExp(`\\b(?:only|just)\\s+(?:the\\s+)?${NAME_EN}\\s+(${NOUN_EN})\\b`, 'gi'), needsExclusive: false },
  // "the Abstract paragraph only"
  { re: new RegExp(`\\b(?:the\\s+)?${NAME_EN}\\s+(${NOUN_EN})\\s+only\\b`, 'gi'), needsExclusive: false },
  // "fix the grammar in the Methods section" — a location, but also how
  // "especially in the Methods section" reads, so it needs a restriction cue.
  { re: new RegExp(`\\b(?:in|within|inside|under)\\s+the\\s+${NAME_EN}\\s+(${NOUN_EN})\\b`, 'gi'), needsExclusive: true },
  // Canonical section names without a noun: "shorten the abstract", "only the conclusion"
  { re: new RegExp(`\\b(?:${EDIT_VERB_EN}|only|just)\\s+(?:only\\s+|just\\s+)?(?:the\\s+)?(${SECTION_NAMES_EN})\\b(?!\\s+${NOUN_EN})`, 'gi'), needsExclusive: false, bare: true },
];

const ZH_NAME = '[“"「『‘\']?([\\u4e00-\\u9fffA-Za-z0-9·]{1,12}?)[”"」』’\']?';
const ZH_NOUN = '(?:这一段|这段|该段|一段|这一部分|这部分|该部分|部分|章节|一节|这一节|这一章|一章|段落|段)';
const ZH_VERB = '(?:修改|改写|重写|润色|优化|精简|缩写|扩写|翻译|调整|完善|改进|修订|校对|压缩|简化|删除|删掉)';
const ZH_SECTION_NAMES = '(?:摘要|引言|前言|绪论|方法|结果|讨论|结论|总结|致谢|背景|概述|概要)';

const TARGET_PATTERNS_ZH = [
  { re: new RegExp(`(?:只|仅|仅仅)?(?:需要|要)?${ZH_VERB}(?:一下)?(?:一下)?${ZH_NAME}(${ZH_NOUN})`, 'g'), needsExclusive: false },
  { re: new RegExp(`(?:只|仅|仅仅)(?:需要|要)?(?:${ZH_VERB})?(?:一下)?${ZH_NAME}(${ZH_NOUN})`, 'g'), needsExclusive: false },
  { re: new RegExp(`(?:只|仅|仅仅|${ZH_VERB})(?:需要|要)?(?:${ZH_VERB})?(?:一下)?(?:“|"|「)?(${ZH_SECTION_NAMES})(?:”|"|」)?`, 'g'), needsExclusive: false, bare: true },
];

/** Positional references that need a human-style paragraph count. */
const POSITIONAL_PATTERNS = [
  /\b(?:the\s+)?(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|last|final|penultimate|\d+(?:st|nd|rd|th))\s+(?:paragraph|section|chapter|part|passage)s?\b/i,
  /\b(?:paragraph|para|section|chapter)\s+#?\d+\b/i,
  /第[一二三四五六七八九十百零两\d]+\s*(?:段|章|节|部分)/,
  /(?:最后一?段|末段|首段|开头一?段|结尾一?段)/,
];

/** Cues that every OTHER part of the document must stay untouched. */
const EXCLUSIVE_PATTERNS = [
  /\bonly\b/i,
  /\b(?:do\s+not|don['’]t|never)\s+(?:change|modify|touch|edit|alter|rewrite|revise|reword)\s+(?:any\s+)?(?:other|another|else|the\s+rest|anything\s+else|remaining)/i,
  /\b(?:leave|keep)\s+(?:all\s+)?(?:the\s+)?(?:rest|everything\s+else|(?:all\s+)?other\s+\w+|remaining\s+\w+)\s+(?:of\s+[\w\s]{1,30}?\s+)?(?:unchanged|alone|untouched|as[- ]is|intact)/i,
  /\b(?:change|edit|touch|modify)\s+nothing\s+else\b|\bnothing\s+else\b/i,
  /\bwithout\s+(?:changing|touching|modifying|editing)\s+(?:anything|any)\s+(?:else|other)/i,
  /(?:只|仅|仅仅)/,
  /其(?:他|余)(?:的)?(?:段落|部分|内容|章节)?(?:都)?(?:不要|不得|请勿|勿|均不|保持不变|保持原样)/,
  /(?:不要|请勿|勿|不得)(?:改动|修改|更改|变动|动|碰)其(?:他|余)/,
];

/** The user explicitly wants the whole document processed. */
const GLOBAL_PATTERNS = [
  /\b(?:all|every|each|entire|whole)\s+(?:the\s+)?(?:paragraphs?|sections?|chapters?|parts?|document|article|text|paper)\b/i,
  /\bthroughout\b|\bdocument[- ]wide\b|\bacross\s+the\s+(?:document|article|paper)\b/i,
  /全文|整篇|整个文档|整个文章|整份文档|所有(?:的)?(?:段落|章节|部分)|每(?:一)?(?:段|章|节)/,
];

/**
 * Reads an instruction for a scope restriction: does it name a specific part
 * of the document and ask for that part only?
 *
 * Deliberately conservative — a missed restriction leaves the existing
 * behaviour (plus the reply guards), while a false one narrows or reroutes a
 * run the user meant to be document-wide. A named target is recognised as the
 * object of an edit verb ("rewrite the Abstract paragraph"), after "only"
 * ("only the Methods section"), or — when the instruction also says to leave
 * everything else alone — as a location ("in the Methods section"). A
 * document-wide cue ("the whole document", "全文") disables the restriction.
 *
 * @param {string} instruction - The user's edit instruction
 * @returns {EditScope}
 */
export function parseEditScope(instruction) {
  const raw = String(instruction || '');
  const none = { restricted: false, targets: [], positional: false };
  if (!raw.trim()) return none;

  // Negated and preservation clauses ("Do not change any other paragraph")
  // are constraints, not objects of the request: read targets from what is left.
  const action = intentText(raw) || raw;
  if (GLOBAL_PATTERNS.some((pattern) => pattern.test(action))) return none;
  const exclusive = EXCLUSIVE_PATTERNS.some((pattern) => pattern.test(raw));

  const targets = [];
  const seen = new Set();
  const add = (name, noun) => {
    const cleaned = String(name || '').trim().replace(/\s+/g, ' ');
    const key = _normalizeLabel(cleaned);
    if (!key || seen.has(key + '|' + noun)) return;
    const words = key.split(' ');
    if (words.some((word) => GENERIC_NAMES.has(word)) && !/[一-鿿]/.test(key)) return;
    if (/^(?:这|该|其他|其余|所有|每|全)/.test(key)) return;
    seen.add(key + '|' + noun);
    targets.push({ name: cleaned, noun });
  };

  for (const { re, needsExclusive, bare } of [...TARGET_PATTERNS_EN, ...TARGET_PATTERNS_ZH]) {
    if (needsExclusive && !exclusive) continue;
    for (const match of action.matchAll(re)) {
      if (bare) add(match[1], 'section');
      else add(match[1], _canonicalNoun(match[2]));
    }
  }
  const positional = POSITIONAL_PATTERNS.some((pattern) => pattern.test(action));

  return { restricted: targets.length > 0 || positional, targets, positional };
}

/** Maps a matched noun to the extent it names: a heading line, or a section body. */
function _canonicalNoun(noun) {
  const lower = String(noun || '').toLowerCase();
  if (lower === 'heading') return 'heading';
  if (/(?:paragraph|passage|portion|段)/.test(lower) && !/部分/.test(lower)) return 'paragraph';
  return 'section';
}

/**
 * Normalises a heading or target name for comparison: lower case, no leading
 * outline numbering ("2.1", "IV.", "第三章"), no surrounding quotes or trailing
 * colon.
 *
 * @param {string} text
 * @returns {string}
 */
function _normalizeLabel(text) {
  let value = String(text || '').toLowerCase().trim();
  value = value.replace(/^["“”‘’'「『]+|["“”‘’'」』]+$/g, '');
  value = value.replace(/^第[一二三四五六七八九十百零两\d]+\s*[章节部分篇]\s*/, '');
  value = value.replace(/^[一二三四五六七八九十]+\s*[、.．]\s*/, '');
  value = value.replace(/^(?:(?:section|chapter|part)\s+)?(?:\d+(?:\.\d+)*|[ivxlcdm]+|[a-z])\s*[.)、:：-]\s*/, '');
  value = value.replace(/^\d+(?:\.\d+)*\s+/, '');
  value = value.replace(/[:：.。\s]+$/g, '');
  return value.replace(/\s+/g, ' ').trim();
}

/** Whether a (normalised) heading names the (normalised) target. */
function _headingMatches(heading, name) {
  if (!heading || !name) return { exact: false, loose: false };
  if (heading === name) return { exact: true, loose: true };
  if (/[一-鿿]/.test(name)) return { exact: false, loose: heading.includes(name) };
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return { exact: false, loose: new RegExp(`(?:^|[^a-z0-9])${escaped}(?:$|[^a-z0-9])`).test(heading) };
}

/**
 * Finds the paragraphs of a parsed document that an {@link EditScope} names.
 *
 * A target resolves to a heading (for `section`/`paragraph` it is the body
 * under that heading, up to the next heading of the same or a higher level; for
 * `heading` the heading line itself), or to a paragraph that opens with the
 * name as a label ("Abstract: ...", "摘要：..."). Resolution fails — `null` —
 * when the scope is positional, when any target is missing, or when a target
 * is ambiguous (several equally good headings), because editing a guessed
 * location is worse than handing the request to a model that can search.
 *
 * @param {Array<{index: number, text: string, headingLevel?: number, inTable?: boolean}>} paragraphs -
 *   Parsed document paragraphs, in document order
 * @param {EditScope} scope
 * @returns {{ indexes: Set<number>, label: string } | null} Document paragraph
 *   indexes that may change, or null when the scope cannot be resolved
 */
export function locateScopeTargets(paragraphs, scope) {
  if (!scope?.restricted || scope.positional || !scope.targets.length) return null;
  const body = Array.isArray(paragraphs) ? paragraphs : [];
  const indexes = new Set();

  for (const target of scope.targets) {
    const name = _normalizeLabel(target.name);
    if (!name) return null;

    const exactHeadings = [];
    const looseHeadings = [];
    body.forEach((para, position) => {
      if (para.inTable || !(para.headingLevel > 0)) return;
      const match = _headingMatches(_normalizeLabel(para.text), name);
      if (match.exact) exactHeadings.push(position);
      else if (match.loose) looseHeadings.push(position);
    });
    const headings = exactHeadings.length ? exactHeadings : looseHeadings;
    if (headings.length > 1) return null;

    if (headings.length === 1) {
      const start = headings[0];
      if (target.noun === 'heading') {
        indexes.add(body[start].index);
        continue;
      }
      const level = body[start].headingLevel;
      let end = start + 1;
      while (end < body.length && !(body[end].headingLevel > 0 && body[end].headingLevel <= level)) end++;
      let found = 0;
      for (let k = start + 1; k < end; k++) {
        if (body[k].inTable || body[k].headingLevel > 0) continue;
        indexes.add(body[k].index);
        found++;
      }
      if (!found) return null;
      continue;
    }

    // No heading: a paragraph that opens with the name as its label.
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const labelRe = new RegExp(`^\\s*["“「]?${escaped}["”」]?\\s*[:：.\\-—–]`, 'i');
    const labelled = body.filter((para) => !para.inTable && !(para.headingLevel > 0) && labelRe.test(para.text || ''));
    if (labelled.length !== 1) return null;
    indexes.add(labelled[0].index);
  }

  return { indexes, label: scope.targets.map((target) => target.name).join(', ') };
}

/**
 * Projects a model reply onto the paragraphs that may change, so everything
 * else stays exactly as it was whatever the model returned.
 *
 * The reply is aligned with the original as the reassembler would; edits to
 * target paragraphs are kept, edits to any other paragraph are reverted, and a
 * paragraph the model inserted is kept only when it follows a target.
 *
 * @param {string[]} origTexts - Original paragraph texts of the chunk
 * @param {boolean[]} isTarget - Per original paragraph: may it change?
 * @param {string} amendedText - The model's (cleaned) reply
 * @returns {{ text: string|null, revertedParagraphs: number }} The confined
 *   amendment text, or null when nothing in scope changed; and how many
 *   out-of-scope paragraph edits (changes, deletions, insertions) were reverted
 */
export function confineAmendment(origTexts, isTarget, amendedText) {
  const orig = origTexts.map((text) => String(text || ''));
  const amended = amendedLines(amendedText);
  const ops = alignParagraphs(orig, amended);
  const out = [];
  let reverted = 0;
  let anchor = -1; // original paragraph most recently placed in the output order

  for (const op of ops) {
    if (op.type === 'keep') {
      anchor = op.origIdx;
      if (isTarget[op.origIdx]) out.push(amended[op.newIdx]);
      else {
        if (orig[op.origIdx].trim() !== amended[op.newIdx].trim()) reverted++;
        out.push(orig[op.origIdx]);
      }
    } else if (op.type === 'delete') {
      anchor = op.origIdx;
      if (isTarget[op.origIdx]) continue;
      reverted++;
      out.push(orig[op.origIdx]);
    } else if (anchor >= 0 && isTarget[anchor]) {
      out.push(amended[op.newIdx]);
    } else if (amended[op.newIdx].trim()) {
      reverted++;
    }
  }

  const changed = out.length !== orig.length || out.some((line, i) => line.trim() !== orig[i].trim());
  return { text: changed ? out.join('\n') : null, revertedParagraphs: reverted };
}

/**
 * Applies a resolved scope to a run's chunk results, in place: every
 * fulfilled amendment is projected onto the target paragraphs with
 * {@link confineAmendment}. A chunk without a target keeps none of its edits,
 * so its amendment becomes null and the chunk is no longer proposed.
 *
 * @param {Array<{status: string, amendment: string|null, chunk?: {paragraphs?: Array<{index: number, text: string}>},
 *   scope?: {label: string, revertedParagraphs: number}}>} results -
 *   ChunkResults from processChunksParallel
 * @param {{ indexes: Set<number>, label: string }} targets - From {@link locateScopeTargets}
 * @returns {{ label: string, revertedParagraphs: number, droppedChunks: number }}
 *   How many out-of-scope paragraph edits were reverted, and how many chunks
 *   lost every edit because they hold no target
 */
export function confineChunkResults(results, targets) {
  let revertedParagraphs = 0;
  let droppedChunks = 0;
  for (const result of results) {
    if (result.status !== 'fulfilled' || typeof result.amendment !== 'string' || !result.amendment) continue;
    const paragraphs = result.chunk?.paragraphs || [];
    const isTarget = paragraphs.map((paragraph) => targets.indexes.has(paragraph.index));
    const confined = confineAmendment(paragraphs.map((paragraph) => paragraph.text || ''), isTarget, result.amendment);
    revertedParagraphs += confined.revertedParagraphs;
    if (confined.text === null && confined.revertedParagraphs > 0) droppedChunks++;
    result.amendment = confined.text;
    result.scope = { label: targets.label, revertedParagraphs: confined.revertedParagraphs };
  }
  return { label: targets.label, revertedParagraphs, droppedChunks };
}

/**
 * Records, in place, how many paragraphs each fulfilled amendment would
 * delete, so the proposal card can state deletions a section-sized inline
 * diff hides.
 *
 * @param {Array<{status: string, amendment: string|null, chunk?: {paragraphs?: Array<{text: string}>},
 *   deletedParagraphs?: number}>} results - ChunkResults
 */
export function markParagraphDeletions(results) {
  for (const result of results) {
    if (result.status !== 'fulfilled' || !result.amendment) continue;
    result.deletedParagraphs = summarizeAmendment((result.chunk?.paragraphs || []).map((paragraph) => paragraph.text || ''), result.amendment).deleted;
  }
}

/**
 * One call for a staged amendment run: resolve the scope the instruction
 * names (or reuse `targets` from the original run, for retries), confine the
 * results to it in place, and record paragraph deletions.
 *
 * @param {Array<object>} results - ChunkResults, modified in place
 * @param {{ paragraphs?: Array<{index: number, text: string, headingLevel?: number, inTable?: boolean}>,
 *   instruction?: string, targets?: { indexes: Set<number>, label: string } | null }} options
 * @returns {{ scope: ({ label: string, notice: string, unresolved?: boolean,
 *   revertedParagraphs?: number, droppedChunks?: number }) | null,
 *   targets: { indexes: Set<number>, label: string } | null }}
 *   `scope` describes the limit for the proposal card (null when the request
 *   names no part); `targets` lets a retry reuse the resolved scope
 */
export function applyEditScope(results, { paragraphs = [], instruction = '', targets = null } = {}) {
  let resolved = targets;
  let scope = null;
  if (!resolved) {
    const editScope = parseEditScope(instruction);
    if (editScope.restricted) {
      resolved = locateScopeTargets(paragraphs, editScope);
      if (!resolved) {
        const label = editScope.targets.map((target) => target.name).join(', ') || 'the named part';
        scope = { label, unresolved: true,
          notice: `This request names ${label}, but it could not be located unambiguously. Every section below may change; review each one.` };
      }
    }
  }
  if (resolved) {
    const confined = confineChunkResults(results, resolved);
    scope = { ...confined, notice: `Limited to ${confined.label}.`
      + (confined.revertedParagraphs ? ` ${confined.revertedParagraphs} model edit(s) outside it were discarded.` : '') };
  }
  markParagraphDeletions(results);
  return { scope, targets: resolved };
}
