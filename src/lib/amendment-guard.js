/**
 * Amendment Guard Module
 *
 * Safeguards for the chunked whole-document amendment pipeline. That pipeline
 * sends EVERY chunk of the document to the model with the user's instruction
 * as the amendment template, then diffs each reply against the chunk. Three
 * failure modes follow from it, none of which the length guard in the
 * reassembler can see:
 *
 *  1. Chat replies accepted as document text. A chunk that does not contain
 *     what the instruction talks about ("Rewrite only the Abstract") gets
 *     answered with an apology ("The Abstract was not included in the text
 *     provided ..."). The reply is long enough to pass the truncation guard, so
 *     it was aligned as "delete every paragraph, insert one paragraph of chat".
 *     {@link classifyChunkReply} detects such replies.
 *
 *  2. No scope control. "Rewrite only X, change nothing else" still rewrites
 *     whatever else the model feels like touching in every other chunk.
 *     The scope guard lives in ./edit-scope.js, which is loaded only when a
 *     staged amendment run starts.
 *
 *  3. Invisible destruction. Replacing or deleting most of a chunk is hard to
 *     see in a chunk-sized inline diff. {@link summarizeAmendment} counts the
 *     paragraph-level deletions so the proposal card can surface them.
 *
 * Pure JavaScript -- no Word API dependency.
 *
 * @module amendment-guard
 */

import { alignParagraphs, amendmentParagraphTexts, paragraphSimilarity, SIMILARITY_THRESHOLD } from './paragraph-alignment.js';
import { textEditHunks, touchesProtectedSpan } from './text-edit-plan.js';
import { preserveNonBreakingSpaces } from './nbsp-preservation.js';
export { protectChunkText } from './protected-chunk-text.js';
export { requestsNbspChange } from './nbsp-preservation.js';

/**
 * @typedef {Object} ReplyVerdict
 * @property {boolean} metaReply - True when the reply is chat, not document text
 * @property {string} reason - Human-readable evidence (empty when not a meta reply)
 */

/**
 * @typedef {Object} AmendmentSummary
 * @property {number} total - Original paragraphs
 * @property {number} unchanged - Paragraphs carried over verbatim
 * @property {number} modified - Paragraphs kept with edited text
 * @property {number} deleted - Original paragraphs with no counterpart in the reply
 * @property {number} inserted - Reply paragraphs with no counterpart in the original
 * @property {boolean} wholesale - Most of the chunk is deleted or replaced
 */

// ---------------------------------------------------------------------------
// Shared text helpers
// ---------------------------------------------------------------------------

/** Non-blank trimmed lines, CR/CRLF normalised. */
function _lines(text) {
  return String(text || '')
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean);
}

/** Lower-cases and collapses whitespace for loose containment checks. */
function _squash(text) {
  return String(text || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// ---------------------------------------------------------------------------
// 1. Meta / refusal replies
// ---------------------------------------------------------------------------

/**
 * Phrases that talk ABOUT the text instead of being text. A single phrase is
 * never enough on its own (a document may say "the text provided"); the
 * classifier also needs the reply to have abandoned the chunk's paragraphs.
 */
const META_PATTERNS = [
  // "the text provided", "the section supplied for amendment", "provided text"
  /\b(?:text|section|passage|excerpt|content|document|paragraph|portion|material|selection)\s+(?:was\s+|were\s+|is\s+|are\s+|has\s+been\s+|have\s+been\s+)?(?:provided|supplied|given|shared|included|presented|submitted)\b/i,
  /\b(?:provided|supplied|given|submitted)\s+(?:text|section|passage|excerpt|content|document|material)\b/i,
  // "was not included", "does not contain", "is not present", "no mention of"
  /\b(?:was|were|is|are)\s+(?:not|n't)\s+(?:included|found|present|contained|part of|available|mentioned|visible|there)\b/i,
  /\b(?:does|do|did)\s+not\s+(?:contain|include|have|appear|mention|feature)\b/i,
  /\b(?:cannot|can't|can not|unable to|couldn't|could not)\s+(?:find|locate|see|identify|edit|amend|rewrite|revise|modify|comply|fulfil|fulfill|complete|perform|apply)\b/i,
  /^(?:I(?:'m| am)\s+(?:sorry|unable|afraid|not able)|I\s+(?:cannot|can't|could not|couldn't|apologi[sz]e|don't see|do not see|didn't find|did not find|was unable|am unable)|sorry\b|my apologies|apologies\b|as an ai\b)/i,
  /\b(?:no|nothing)\s+(?:changes?|edits?|amendments?|modifications?|revisions?)\s+(?:were|was|are|is|needed|required|necessary|to\s+(?:be\s+)?made)\b/i,
  /\bplease\s+(?:provide|share|send|supply|include|paste|upload|clarify)\b/i,
  // Chinese
  /(?:提供|给出|附上|发来)的(?:文本|内容|部分|段落|文档|文字|片段)/,
  /(?:未|没有|并未|不)(?:包含|找到|提供|出现|涉及|存在)/,
  /无法(?:找到|修改|完成|执行|处理|定位|按照)/,
  /(?:抱歉|对不起|很遗憾)/,
  /作为(?:一个)?(?:AI|人工智能)/i,
  /(?:无需|不需要|没有必要)(?:修改|更改|改动|变动)/,
  /请(?:提供|补充|告知|说明)/,
];

/**
 * Minimum share of the chunk's paragraphs that must still be recognisable in
 * a reply for the reply to count as an edit of that chunk. A genuine
 * amendment carries nearly every paragraph over (edited or not).
 */
const MIN_CARRIED_FRACTION = 0.5;

/**
 * Decides whether a model reply for a chunk is chat about the chunk rather
 * than the chunk's text.
 *
 * A reply is a meta reply when BOTH hold:
 *   - it contains lines that are new (absent from the chunk) and read as
 *     discourse about the text ("not included in the text provided",
 *     "I cannot find", "no changes were needed", "无法找到…"), and
 *   - fewer than half of the chunk's paragraphs survive in it, i.e. it did not
 *     edit the chunk — it replaced it.
 *
 * Lines the user's own instruction quotes are never counted as chat, so
 * "replace this section with: The text provided is withdrawn" still works.
 * An edit that keeps the chunk but adds a note is NOT dropped: the chunk is
 * recognisable, and the user sees the stray paragraph in the proposal diff.
 *
 * @param {string[]} origTexts - Original paragraph texts of the chunk
 * @param {string} replyText - The model's (cleaned) reply
 * @param {{ instruction?: string }} [options]
 * @returns {ReplyVerdict}
 */
export function classifyChunkReply(origTexts, replyText, { instruction = '' } = {}) {
  const replyLines = _lines(replyText);
  if (replyLines.length === 0) return { metaReply: false, reason: '' };

  const origLines = origTexts.flatMap(_lines);
  if (origLines.length === 0) return { metaReply: false, reason: '' };
  const origSet = new Set(origLines);
  const instructionSquashed = _squash(instruction);

  // Original paragraphs still recognisable in the reply (verbatim or edited).
  let carried = 0;
  for (const orig of origLines) {
    if (replyLines.some((line) => line === orig || paragraphSimilarity(orig, line) >= SIMILARITY_THRESHOLD)) carried++;
  }
  const carriedFraction = carried / origLines.length;
  if (carriedFraction >= MIN_CARRIED_FRACTION) return { metaReply: false, reason: '' };

  for (const line of replyLines) {
    if (origSet.has(line)) continue;
    if (instructionSquashed && instructionSquashed.includes(_squash(line))) continue;
    if (origLines.some((orig) => paragraphSimilarity(orig, line) >= SIMILARITY_THRESHOLD)) continue;
    if (META_PATTERNS.some((pattern) => pattern.test(line))) {
      const quoted = line.length > 120 ? `${line.slice(0, 117)}...` : line;
      return {
        metaReply: true,
        reason: `reply reads as an explanation, not document text ("${quoted}"), and keeps only ${carried} of ${origLines.length} original paragraph(s)`,
      };
    }
  }
  return { metaReply: false, reason: '' };
}

// ---------------------------------------------------------------------------
// 2. Paragraph-level change summary
// ---------------------------------------------------------------------------

/**
 * A reply whose deletions reach this share of the chunk (with at least two
 * paragraphs gone) is flagged as deleting or replacing the chunk wholesale.
 */
const WHOLESALE_DELETE_FRACTION = 0.5;

/**
 * Counts what a reply does to a chunk at paragraph level, using the same
 * alignment the reassembler applies. A paragraph the model rewrote beyond
 * recognition aligns as delete + insert, so a wholesale rewrite and a wholesale
 * deletion both show up as `wholesale`.
 *
 * @param {string[]} origTexts - Original paragraph texts of the chunk
 * @param {string} amendedText - The model's (cleaned) reply
 * @returns {AmendmentSummary}
 */
export function summarizeAmendment(origTexts, amendedText) {
  const orig = origTexts.map((text) => String(text || ''));
  const amended = amendmentParagraphTexts(orig, amendedText);
  const ops = alignParagraphs(orig, amended);
  let unchanged = 0;
  let modified = 0;
  let deleted = 0;
  let inserted = 0;
  for (const op of ops) {
    if (op.type === 'keep') {
      if (orig[op.origIdx].trim() === amended[op.newIdx].trim()) unchanged++;
      else modified++;
    } else if (op.type === 'delete') deleted++;
    else inserted++;
  }
  const total = orig.length;
  return {
    total, unchanged, modified, deleted, inserted,
    wholesale: deleted >= 2 && deleted >= total * WHOLESALE_DELETE_FRACTION,
  };
}

/** Reply lines as the reassembler sees them: no leading/trailing blank lines. */
export function amendedLines(text) {
  const lines = String(text || '').replace(/\r\n?/g, '\n').split('\n');
  while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
  while (lines.length && lines[0].trim() === '') lines.shift();
  return lines;
}

/** Check the same paragraph constraints used by the writer before staging a preview. */
export function checkChunkAmendment(chunk, amendment, { preserveNbsp = false } = {}) {
  const paragraphs = (chunk.paragraphs || []).filter((p) => String(p.text || '').trim());
  const original = paragraphs.map((p) => String(p.text));
  let amended;
  try { amended = amendmentParagraphTexts(original, amendment); }
  catch (error) { return { valid: false, issues: [error.message] }; }
  const issues = [];
  const beforeChars = original.reduce((count, text) => count + text.length, 0);
  const afterChars = amended.reduce((count, text) => count + text.length, 0);
  if (beforeChars && afterChars < beforeChars * 0.3) {
    issues.push('Output appears truncated (less than 30% of the source). Return the complete amended section, including unchanged paragraphs.');
  }
  const alignment = original.length === amended.length && paragraphs.some((p) => p.hasRevisions || p.protectedSpans?.length)
    ? original.map((_text, index) => ({ type: 'keep', origIdx: index, newIdx: index }))
    : alignParagraphs(original, amended);
  if (preserveNbsp) {
    for (const op of alignment) {
      if (op.type === 'keep') amended[op.newIdx] = preserveNonBreakingSpaces(original[op.origIdx], amended[op.newIdx]);
    }
  }
  let anchor = -1;
  for (const op of alignment) {
    if (op.type === 'delete') {
      const paragraph = paragraphs[op.origIdx];
      if (paragraph.hasRevisions || paragraph.inTable) {
        issues.push(`Paragraph ${op.origIdx + 1} cannot be deleted or replaced structurally (${paragraph.inTable ? 'table content' : 'unresolved revisions'}). Preserve it and edit its text in place.`);
      }
      if (paragraph.protectedSpans?.length) issues.push(`Paragraph ${op.origIdx + 1} contains native fields or equations. Keep this paragraph and its protected content intact.`);
    } else if (op.type === 'keep' && paragraphs[op.origIdx].protectedSpans?.length
        && textEditHunks(original[op.origIdx], amended[op.newIdx], paragraphs[op.origIdx].protectedSpans)
          .some((hunk) => touchesProtectedSpan(hunk, paragraphs[op.origIdx].protectedSpans))) {
      issues.push(`Paragraph ${op.origIdx + 1} changes a native field result or equation. Preserve its exact formula, citation, bibliography or link text and edit surrounding prose only.`);
    } else if (op.type === 'insert' && amended[op.newIdx].trim() && paragraphs[Math.max(0, anchor)]?.inTable) {
      issues.push(`Cannot insert a prose paragraph inside table content near paragraph ${Math.max(0, anchor) + 1}. Preserve the table paragraph structure.`);
    }
    if (op.origIdx !== undefined) anchor = op.origIdx;
  }
  return { valid: issues.length === 0, issues, amendment: amended.join('\n') };
}
