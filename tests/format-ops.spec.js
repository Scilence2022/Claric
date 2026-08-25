/**
 * Format ops tests: the JSON contract between the LLM and the Word format
 * applier. parseFormatOps must tolerate fences/prose and enforce the
 * allowlist so nothing arbitrary reaches the Word API.
 */

const { buildFormatPrompt, parseFormatOps } = require('../src/lib/format-ops.js');

describe('parseFormatOps', () => {
  test('parses a bare JSON array', () => {
    const ops = parseFormatOps('[{"font":{"bold":true}}]');
    expect(ops).toEqual([{ font: { bold: true } }]);
  });

  test('strips code fences', () => {
    const ops = parseFormatOps('```json\n[{"font":{"italic":true}}]\n```');
    expect(ops).toEqual([{ font: { italic: true } }]);
  });

  test('tolerates surrounding prose', () => {
    const ops = parseFormatOps('Sure! Here are the ops:\n[{"font":{"bold":true}}]\nHope that helps.');
    expect(ops).toEqual([{ font: { bold: true } }]);
  });

  test('invalid JSON returns [] with a warning', () => {
    const log = jest.fn();
    expect(parseFormatOps('[{oops}]', log)).toEqual([]);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('not valid JSON'), 'warning');
  });

  test('no JSON array returns [] with a warning', () => {
    const log = jest.fn();
    expect(parseFormatOps('I cannot help with that.', log)).toEqual([]);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('no JSON array'), 'warning');
  });

  test('empty / nullish input returns []', () => {
    expect(parseFormatOps('')).toEqual([]);
    expect(parseFormatOps(null)).toEqual([]);
    expect(parseFormatOps(undefined)).toEqual([]);
  });

  test('non-array JSON returns []', () => {
    expect(parseFormatOps('{"font":{"bold":true}}')).toEqual([]);
  });

  test('entries without a font/paragraph payload are dropped', () => {
    const log = jest.fn();
    const ops = parseFormatOps('[{"match":"hello"},{"font":{"bold":true}}]', log);
    expect(ops).toEqual([{ font: { bold: true } }]);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('no valid font/paragraph payload'), 'warning');
  });

  test('unknown font properties are dropped, known ones kept', () => {
    const ops = parseFormatOps('[{"font":{"bold":true,"evil":"x","color":"#ff0000"}}]');
    expect(ops).toEqual([{ font: { bold: true, color: '#FF0000' } }]);
  });

  test('invalid color is dropped with a warning; op survives on other keys', () => {
    const log = jest.fn();
    const ops = parseFormatOps('[{"font":{"color":"red","bold":true}}]', log);
    expect(ops).toEqual([{ font: { bold: true } }]);
    expect(log).toHaveBeenCalledWith(expect.stringContaining('invalid color'), 'warning');
  });

  test('font with only an invalid color leaves no payload — op dropped', () => {
    expect(parseFormatOps('[{"font":{"color":"red"}}]')).toEqual([]);
  });

  test('match and paragraphStyle selectors are preserved; match wins when both given', () => {
    const ops = parseFormatOps('[{"match":"exact text","font":{"bold":true}},{"paragraphStyle":"heading1","paragraph":{"alignment":"centered"}}]');
    expect(ops[0]).toEqual({ match: 'exact text', font: { bold: true } });
    expect(ops[1]).toEqual({ paragraphStyle: 'heading1', paragraph: { alignment: 'centered' } });

    const both = parseFormatOps('[{"match":"m","paragraphStyle":"heading2","font":{"italic":true}}]');
    expect(both).toEqual([{ match: 'm', font: { italic: true } }]);
  });

  test('numeric strings are coerced; out-of-range numbers are dropped', () => {
    const ops = parseFormatOps('[{"font":{"size":"14"},"paragraph":{"leftIndent":"18","spaceBefore":99999}}]');
    expect(ops).toEqual([{ font: { size: 14 }, paragraph: { leftIndent: 18 } }]);
  });

  test('boolean coercion and enum string trimming', () => {
    const ops = parseFormatOps('[{"font":{"bold":1,"underline":" single ","highlightColor":"yellow"}}]');
    expect(ops).toEqual([{ font: { bold: true, underline: 'single', highlightColor: 'yellow' } }]);
  });

  test('style + styleBuiltIn keeps styleBuiltIn', () => {
    const log = jest.fn();
    const ops = parseFormatOps('[{"paragraph":{"style":"My Style","styleBuiltIn":"heading2"}}]', log);
    expect(ops).toEqual([{ paragraph: { styleBuiltIn: 'heading2' } }]);
  });

  test('negative indents within range are allowed', () => {
    const ops = parseFormatOps('[{"paragraph":{"firstLineIndent":-18}}]');
    expect(ops).toEqual([{ paragraph: { firstLineIndent: -18 } }]);
  });

  test('non-object entries are dropped', () => {
    const ops = parseFormatOps('[null,"x",42,[{"font":{"bold":true}}],{"font":{"bold":true}}]');
    expect(ops).toEqual([{ font: { bold: true } }]);
  });
});

describe('buildFormatPrompt', () => {
  test('embeds the instruction, scope text and scope name', () => {
    const p = buildFormatPrompt('把标题加粗', '第一段\n第二段', 'document');
    expect(p).toContain('把标题加粗');
    expect(p).toContain('第一段\n第二段');
    expect(p).toContain('DOCUMENT TEXT');
    expect(p).toContain('the entire document');
  });

  test('selection scope is named accordingly', () => {
    const p = buildFormatPrompt('make it bold', 'selected words', 'selection');
    expect(p).toContain('SELECTION TEXT');
    expect(p).toContain('the entire selection');
  });

  test('states the JSON-only output contract', () => {
    const p = buildFormatPrompt('x', 'y', 'selection');
    expect(p).toContain('Output ONLY a JSON array');
    expect(p).toContain('output exactly []');
  });
});
