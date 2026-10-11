/**
 * Shared normalization for the heuristic intent detectors.
 *
 * The routing regexes test the user's wording, and two phrasings defeat a
 * raw-text test:
 *  - constraint clauses ("polish it, but don't change the formatting",
 *    "不要修改文档，只告诉我字数") name actions the user does NOT want, yet
 *    their verbs and nouns still match the edit/format detectors;
 *  - polite requests ("Can you fix the typo?", "你能帮我润色这段话吗？")
 *    start like questions, so a question check sends them to Q&A.
 *
 * Every detector normalizes through this module so they agree on what the
 * requested action is. Pipelines still receive the user's full instruction,
 * constraints included. No regex lookbehind: older Safari-based Word hosts
 * reject it at parse time, which would fail the whole module.
 *
 * Pure module.
 *
 * @module intent-text
 */

/** Clause boundaries (captured so the replacement keeps them). */
const EN_BOUNDARY = '(^|[\\s,;:.!?，；：。！？—–])';
const ZH_BOUNDARY = '(^|[\\s，,;；：:。！!？?、]|但是?|而且|并且|且|同时|也|并|请)';
const CLAUSE_TAIL_EN = '[^,;:.!?，；：。！？—–]*';
const CLAUSE_TAIL_ZH = '[^，,;；：:。！!？?]*';

/**
 * Negated actions are never the requested action. "Don't forget to …" is
 * an instruction, so forget/忘 is excluded, and 不要的 ("unwanted") is an
 * adjective, not a negation.
 */
const EN_NEGATED_CLAUSE_RE = new RegExp(`${EN_BOUNDARY}(?:(?:but|and|yet|however)\\s*,?\\s+)?(?:please\\s+)?`
    + `(?:(?:do\\s+not|don['’]t|dont|never|no\\s+need\\s+to)(?!\\s+forget)\\b`
    + `|without\\s+(?:changing|altering|modifying|touching|affecting|editing|rewriting|removing|deleting))${CLAUSE_TAIL_EN}`, 'gi');
const ZH_NEGATED_CLAUSE_RE = new RegExp(`${ZH_BOUNDARY}(?:请)?`
    + `(?:(?:不要|不用|无需|不必|不需要|别|切勿|勿|禁止|不得)(?![忘的])|不(?:改变|改动|修改|影响|调整|变动|删除|增加))${CLAUSE_TAIL_ZH}`, 'g');

/**
 * Preservation clauses that only protect formatting/structure. They carry
 * format nouns but request no formatting change. A preservation clause that
 * is the whole input ("保持标题居中") is the instruction itself, so these are
 * stripped only when another clause remains.
 */
const EN_PRESERVE_CLAUSE_RE = new RegExp(`${EN_BOUNDARY}(?:(?:but|and|while)\\s+)?(?:please\\s+)?`
    + `(?:keep|preserve|retain|maintain|leave)\\b[^,;:.!?，；：。！？—–]*?`
    + `\\b(?:formatting|formats?|styles?|headings?|layout|structure|fonts?)\\b${CLAUSE_TAIL_EN}`, 'gi');
const ZH_PRESERVE_CLAUSE_RE = new RegExp(`${ZH_BOUNDARY}(?:请)?(?:保留|保持|维持|沿用)${CLAUSE_TAIL_ZH}`
    + '?(?:格式|标题|样式|排版|版式|结构|字体)' + CLAUSE_TAIL_ZH, 'g');

/** Limiters left at the front once a constraint clause is gone ("…, just list the dates"). */
const LEADING_LIMITER_RE = /^(?:(?:just|only|simply|instead)\s+|(?:只需要?|只要|而是|只|仅|就)(?=\S))/i;
const EDGE_NOISE_RE = /^[\s,;:，；：、—–-]+|[\s,;:，；：、—–-]+$/g;
// A connector orphaned by the removal ("请润色但不要改变格式" → "请润色但").
const TRAILING_CONNECTOR_RE = /(?:\s+(?:but|and|yet|while)|但是?|而且|并且|同时)\s*$/i;

const EN_POLITE_RE = /^(?:please\s+)?(?:can|could|would|will)\s+you\s+(?:(?:please|kindly|be\s+able\s+to)\s+)*/i;
// Chinese polite requests need an explicit 帮(我|忙) / 麻烦 marker; a bare
// "你能…吗？" stays a capability question ("你能改变选择的表格的样式吗？").
const ZH_POLITE_RE = /^(?:(?:你|您)?(?:能否|能不能|能|可以|可否)(?:帮我|帮忙)|麻烦(?:你|您)?(?:帮我|帮忙)?|请(?:你|您)?帮(?:我|忙))/;
const POLITE_TAIL_RE = /[\s,，]*(?:please)?\s*(?:好吗|可以吗|行吗|吗|呢)?[\s?？。.!！]*$/i;
const INTERROGATIVE_RE = /\b(?:what|which|who|whom|whose|why|how|whether|when|where)\b|什么|哪些|哪个|哪里|为什么|为何|如何|怎么|怎样|多少|是否|是不是|有没有/i;
const NOT_A_REQUEST_RE = /^(?:like|mind)\b/i;
const PLEASE_RE = /^(?:(?:please|kindly)\s+|请(?!问))/i;

function stripClauses(text, regexes) {
    return regexes.reduce((current, re) => current.replace(re, '$1'), text)
        .replace(EDGE_NOISE_RE, '').replace(TRAILING_CONNECTOR_RE, '').replace(EDGE_NOISE_RE, '').trim();
}

/**
 * Returns the action text the user is asking for: negated clauses removed,
 * formatting-preservation clauses removed (unless that leaves nothing), and a
 * leading "just/only/只" limiter dropped. An input that only negates actions
 * ("Don't edit anything") yields ''.
 *
 * @param {string} text - Raw chat input
 * @returns {string}
 */
export function intentText(text) {
    const trimmed = String(text || '').trim();
    const withoutNegations = stripClauses(trimmed, [EN_NEGATED_CLAUSE_RE, ZH_NEGATED_CLAUSE_RE]);
    const withoutPreservation = stripClauses(withoutNegations, [EN_PRESERVE_CLAUSE_RE, ZH_PRESERVE_CLAUSE_RE]);
    const action = withoutPreservation || withoutNegations;
    return action.replace(LEADING_LIMITER_RE, '').trim();
}

/**
 * Body of a polite request ("Can you fix the typo?" → "fix the typo";
 * "你能帮我润色这段话吗？" → "润色这段话"). Returns null when the text is
 * not a polite request, or when the body still asks something (an
 * interrogative word: "Can you tell me who…", "…有什么风险吗"), so those
 * stay questions.
 *
 * @param {string} text - Normalized action text
 * @returns {string|null}
 */
export function politeRequestBody(text) {
    const trimmed = String(text || '').trim();
    const match = trimmed.match(EN_POLITE_RE) || trimmed.match(ZH_POLITE_RE);
    if (!match) return null;
    const body = trimmed.slice(match[0].length).replace(POLITE_TAIL_RE, '').trim();
    if (!body || NOT_A_REQUEST_RE.test(body) || INTERROGATIVE_RE.test(body)) return null;
    return body;
}

/**
 * Drops a leading "please"/"请" (but not "请问", which asks a question).
 *
 * @param {string} text
 * @returns {string}
 */
export function stripPlease(text) {
    return String(text || '').trim().replace(PLEASE_RE, '').trim();
}
