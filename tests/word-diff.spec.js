/**
 * Word-diff layer tests: the pure diff helpers behind the vendored
 * strategies. The Word-facing apply functions are covered separately
 * (char-diff.spec.js mock harness) or via reassembler.spec.js mocks.
 *
 * Key regression covered here: diff_sentenceMode must diff the
 * occurrence-ordered sentence sequence. Upstream diffed the DEDUPED
 * sentence list, so a repeated sentence silently misaligned every later
 * sentence (wrong deletions/insertions with no error raised).
 */

const { computeDiff, diff_sentenceMode, sliceSearchPieces } = require('../src/lib/word-diff/index.js');

describe('diff_sentenceMode', () => {
  test('repeated sentences keep their positions (one DELETE for the repeat)', () => {
    // "Same. " appears twice in the original, once in the amended text.
    // The correct diff deletes exactly one occurrence — a deduped sequence
    // would report no change at all. (Which occurrence DMP picks is an
    // implementation detail; assert semantics, not the exact alignment.)
    const diffs = diff_sentenceMode('Same. Same. End. ', 'Same. End. ');
    expect(diffs.filter(([op]) => op === -1)).toEqual([[-1, 'Same. ']]);
    expect(diffs.filter(([op]) => op === 1)).toEqual([]);
    // The diff round-trips: EQUAL + INSERT text rebuilds the amended text.
    const rebuilt = diffs.filter(([op]) => op !== -1).map(([, t]) => t).join('');
    expect(rebuilt).toBe('Same. End. ');
  });

  test('inserted sentence appears as one INSERT between EQUALs', () => {
    const diffs = diff_sentenceMode('Intro. End. ', 'Intro. New middle. End. ');
    expect(diffs).toEqual([
      [0, 'Intro. '],
      [1, 'New middle. '],
      [0, 'End. '],
    ]);
  });

  test('identical texts produce a single EQUAL run', () => {
    const diffs = diff_sentenceMode('One. Two. ', 'One. Two. ');
    expect(diffs).toEqual([[0, 'One. Two. ']]);
  });

  test('text without sentence boundaries is one sentence', () => {
    const diffs = diff_sentenceMode('hello', 'hello world');
    expect(diffs).toEqual([[-1, 'hello'], [1, 'hello world']]);
  });
});

describe('computeDiff (word mode)', () => {
  test('a one-word insertion is a single minimal INSERT op', () => {
    const diffs = computeDiff('the quick fox jumps', 'the quick brown fox jumps');
    expect(diffs).toEqual([
      [0, 'the quick '],
      [1, 'brown '],
      [0, 'fox jumps'],
    ]);
  });

  test('identical texts produce a single EQUAL op', () => {
    expect(computeDiff('same words here', 'same words here')).toEqual([[0, 'same words here']]);
  });
});

describe('sliceSearchPieces', () => {
  test('short text stays one piece', () => {
    expect(sliceSearchPieces('short')).toEqual(['short']);
  });

  test('long text splits at the 200-char cap and rejoins losslessly', () => {
    const text = 'x'.repeat(450);
    const pieces = sliceSearchPieces(text);
    expect(pieces.map((p) => p.length)).toEqual([200, 200, 50]);
    expect(pieces.join('')).toBe(text);
  });

  test('a surrogate pair on the boundary is never split', () => {
    // 199 ASCII chars + 😀 (a surrogate pair) + tail: the naive 200-char cut
    // would end the first piece on the high surrogate.
    const text = 'a'.repeat(199) + '😀' + 'b'.repeat(100);
    const pieces = sliceSearchPieces(text);
    expect(pieces.join('')).toBe(text);
    expect(pieces.length).toBe(2);
    expect(pieces[0]).toBe('a'.repeat(199));
    expect(pieces[1]).toBe('😀' + 'b'.repeat(100));
    for (const piece of pieces) {
      const first = piece.charCodeAt(0);
      const last = piece.charCodeAt(piece.length - 1);
      // No lone surrogate halves at the seams.
      expect(first >= 0xdc00 && first <= 0xdfff).toBe(false);
      expect(last >= 0xd800 && last <= 0xdbff).toBe(false);
    }
  });
});
