/**
 * Regression pin for the vendored diff-match-patch library
 * (src/lib/vendor/diff-match-patch.js — verbatim npm diff-match-patch@1.0.5).
 *
 * Captures the exact behaviors a future library swap must preserve:
 *  - diff_cleanupSemantic op shapes (used by ui/diff-view.js),
 *  - the word-mode prototype extension in word-diff/diff-wordmode.js,
 *  - the char-level API behind word-diff/char-diff.js.
 */

const DiffMatchPatch = require('../src/lib/vendor/diff-match-patch.js');
const WordModeDMP = require('../src/lib/word-diff/diff-wordmode.js').default;
const { computeCharEdits } = require('../src/lib/word-diff/char-diff.js');

describe('vendored diff-match-patch', () => {
  test('exposes the constructor and DIFF op constants like the npm package did', () => {
    expect(typeof DiffMatchPatch).toBe('function');
    expect(DiffMatchPatch.DIFF_DELETE).toBe(-1);
    expect(DiffMatchPatch.DIFF_INSERT).toBe(1);
    expect(DiffMatchPatch.DIFF_EQUAL).toBe(0);
  });

  test('diff_cleanupSemantic merges fragmented char ops into word-level ops', () => {
    const dmp = new DiffMatchPatch();
    const diffs = dmp.diff_main(
      'The quick brown fox jumps over the lazy dog.',
      'The quick red fox jumps over the lazy dog.'
    );
    dmp.diff_cleanupSemantic(diffs);
    // Raw diff_main fragments this into [-1,'b'],[0,'r'],[-1,'own'],[1,'ed'];
    // semantic cleanup must realign it to whole-word delete/insert ops.
    expect(diffs).toEqual([
      [0, 'The quick '],
      [-1, 'brown'],
      [1, 'red'],
      [0, ' fox jumps over the lazy dog.'],
    ]);
  });

  test('word-mode diff (prototype extension) produces word-granular ops', () => {
    const dmp = new WordModeDMP();
    expect(dmp.diff_wordMode('the quick fox jumps', 'the quick brown fox jumps')).toEqual([
      [0, 'the quick '],
      [1, 'brown '],
      [0, 'fox jumps'],
    ]);
  });

  test('char-level computeCharEdits keeps a one-comma CJK edit minimal', () => {
    expect(computeCharEdits(
      '但有一个女孩选择不备份自己与母亲的回忆。',
      '但有一个女孩，选择不备份自己与母亲的回忆。'
    )).toEqual([
      [0, '但有一个女孩'],
      [1, '，'],
      [0, '选择不备份自己与母亲的回忆。'],
    ]);
    expect(computeCharEdits('same text', 'same text')).toEqual([[0, 'same text']]);
  });
});
