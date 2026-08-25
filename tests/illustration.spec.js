/** @jest-environment jsdom */

/**
 * Illustration tests: the SVG contract between the LLM and the image
 * inserter. parseIllustration must tolerate fences/prose; sanitizeSvg must
 * strip active content (jsdom is required — DOMPurify needs a DOM) before
 * the markup reaches the preview DOM or Word.
 */

const {
  buildIllustrationPrompt, parseIllustration, sanitizeSvg,
  svgDimensions, ensureSvgDimensions, illustrationPositionFromInstruction,
} = require('../src/lib/illustration.js');

const SIMPLE_SVG = '<svg width="1200" height="800" viewBox="0 0 1200 800"><rect width="1200" height="800" fill="#123"/></svg>';

describe('parseIllustration', () => {
  test('parses a bare SVG document', () => {
    const result = parseIllustration(SIMPLE_SVG);
    expect(result).toEqual({ svg: SIMPLE_SVG });
  });

  test('strips code fences', () => {
    const result = parseIllustration('```svg\n' + SIMPLE_SVG + '\n```');
    expect(result).toEqual({ svg: SIMPLE_SVG });
  });

  test('tolerates surrounding prose, slicing the first svg element', () => {
    const result = parseIllustration('Here is your artwork:\n' + SIMPLE_SVG + '\nHope you like it!');
    expect(result).toEqual({ svg: SIMPLE_SVG });
  });

  test('no svg element returns null with a warning', () => {
    const log = jest.fn();
    expect(parseIllustration('I cannot draw that.', log)).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('no <svg> element'), 'warning');
  });

  test('unclosed svg element returns null', () => {
    expect(parseIllustration('<svg width="10"><rect/></svgOOPS')).toBeNull();
  });

  test('empty / nullish input returns null', () => {
    expect(parseIllustration('')).toBeNull();
    expect(parseIllustration(null)).toBeNull();
    expect(parseIllustration(undefined)).toBeNull();
  });

  test('pathologically large payloads are rejected with a warning', () => {
    const log = jest.fn();
    const huge = `<svg width="10" height="10">${'x'.repeat(300 * 1024)}</svg>`;
    expect(parseIllustration(huge, log)).toBeNull();
    expect(log).toHaveBeenCalledWith(expect.stringContaining('too large'), 'warning');
  });
});

describe('sanitizeSvg', () => {
  test('strips scripts, event handlers and foreignObject', () => {
    const dirty = '<svg width="10" height="10" onload="alert(1)"><script>alert(2)</script>'
      + '<foreignObject><body xmlns="http://www.w3.org/1999/xhtml">x</body></foreignObject>'
      + '<rect width="10" height="10"/></svg>';
    const clean = sanitizeSvg(dirty);
    expect(clean).not.toContain('script');
    expect(clean).not.toContain('onload');
    expect(clean).not.toContain('foreignObject');
    expect(clean).toContain('<rect');
  });

  test('keeps shapes, gradients and filters', () => {
    const rich = '<svg width="10" height="10"><defs><linearGradient id="g"><stop offset="0" stop-color="#fff"/></linearGradient></defs>'
      + '<filter id="f"><feGaussianBlur stdDeviation="2"/></filter>'
      + '<rect width="10" height="10" fill="url(#g)" filter="url(#f)"/></svg>';
    const clean = sanitizeSvg(rich);
    expect(clean).toContain('linearGradient');
    expect(clean).toContain('feGaussianBlur');
    expect(clean).toContain('<rect');
  });
});

describe('svgDimensions', () => {
  test('reads width/height attributes (units ignored)', () => {
    expect(svgDimensions('<svg width="1200px" height="800px"></svg>')).toEqual({ width: 1200, height: 800 });
  });

  test('falls back to the viewBox extent', () => {
    expect(svgDimensions("<svg viewBox='10 10 600 400'></svg>")).toEqual({ width: 600, height: 400 });
  });

  test('returns null when neither is present', () => {
    expect(svgDimensions('<svg><rect/></svg>')).toBeNull();
    expect(svgDimensions('nonsense')).toBeNull();
  });
});

describe('ensureSvgDimensions', () => {
  test('leaves a dimensioned svg untouched', () => {
    expect(ensureSvgDimensions(SIMPLE_SVG)).toBe(SIMPLE_SVG);
  });

  test('injects viewBox-derived dimensions when missing', () => {
    const out = ensureSvgDimensions('<svg viewBox="0 0 600 400"><rect/></svg>');
    expect(out).toContain('<svg width="600" height="400"');
    expect(out).toContain('viewBox="0 0 600 400"');
  });

  test('injects the 1200x800 default when nothing is declared', () => {
    expect(ensureSvgDimensions('<svg><rect/></svg>')).toContain('<svg width="1200" height="800"');
  });
});

describe('illustrationPositionFromInstruction', () => {
  test('hero-image phrasings land at the document start', () => {
    expect(illustrationPositionFromInstruction('给文章加一张题图')).toBe('start');
    expect(illustrationPositionFromInstruction('在开头加一幅插图')).toBe('start');
    expect(illustrationPositionFromInstruction('add an image at the top')).toBe('start');
  });

  test('everything else appends at the end', () => {
    expect(illustrationPositionFromInstruction('设计并增加SVG插图')).toBe('end');
    expect(illustrationPositionFromInstruction('insert an illustration')).toBe('end');
  });
});

describe('buildIllustrationPrompt', () => {
  test('states the SVG-only output contract and embeds the context', () => {
    const p = buildIllustrationPrompt('设计并增加SVG插图', '正文内容');
    expect(p).toContain('设计并增加SVG插图');
    expect(p).toContain('正文内容');
    expect(p).toContain('Output ONLY the SVG markup');
    expect(p).toContain('self-contained');
    expect(p).toContain('viewBox');
  });
});
