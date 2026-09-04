/** @jest-environment jsdom */

/**
 * Illustration tests: the SVG contract between the LLM and the image
 * inserter. parseIllustration must tolerate fences/prose; sanitizeSvg must
 * strip active content (jsdom is required — DOMPurify needs a DOM) before
 * the markup reaches the preview DOM or Word.
 */

const {
  buildIllustrationPrompt, buildIllustrationRedesignPrompt, parseIllustration, sanitizeSvg,
  svgDimensions, ensureSvgDimensions, illustrationPositionFromInstruction,
  illustrationPositionLabel, buildImagePrompt, illustrationRenderer,
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

  test('cursor-anchored phrasings insert at the caret (ZH + EN)', () => {
    expect(illustrationPositionFromInstruction('设计一张简单的图片，然后在光标处插入')).toBe('cursor');
    expect(illustrationPositionFromInstruction('在当前位置插入一张插图')).toBe('cursor');
    expect(illustrationPositionFromInstruction('在此处配一张图')).toBe('cursor');
    expect(illustrationPositionFromInstruction('insert a picture at the cursor')).toBe('cursor');
    expect(illustrationPositionFromInstruction('add an image here')).toBe('cursor');
  });

  test('cursor wins over start when both match', () => {
    expect(illustrationPositionFromInstruction('在光标处插入一张题图')).toBe('cursor');
  });

  test('words merely containing "here"/cursor letters do not match', () => {
    expect(illustrationPositionFromInstruction('设计一张别具匠心的图')).toBe('end');
    expect(illustrationPositionFromInstruction('nowhere to insert')).toBe('end');
  });
});

describe('illustrationPositionLabel', () => {
  test('maps positions to human-readable labels', () => {
    expect(illustrationPositionLabel('start')).toBe('document start');
    expect(illustrationPositionLabel('end')).toBe('document end');
    expect(illustrationPositionLabel('cursor')).toBe('the cursor');
    expect(illustrationPositionLabel(undefined)).toBe('document end');
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

describe('buildIllustrationRedesignPrompt', () => {
  test('keeps the SVG output contract and asks for legible labels', () => {
    const p = buildIllustrationRedesignPrompt('make the legend text larger', '架构正文');
    expect(p).toContain('make the legend text larger');
    expect(p).toContain('架构正文');
    expect(p).toContain('Output ONLY the SVG markup');
    expect(p).toContain('self-contained');
    expect(p).toContain('viewBox');
    // A redesign is a diagram with labels — the opposite of buildImagePrompt,
    // which forbids rendered text.
    expect(p).toContain('text labels are expected');
    expect(p).toContain('no overlaps');
    expect(p).toContain('Keep the original layout');
  });

  test('without a source image, the figure is only the instruction description', () => {
    const p = buildIllustrationRedesignPrompt('redraw the diagram', '正文');
    expect(p).toContain('described in the user instruction');
    expect(p).not.toContain('attached as an image');
  });

  test('with a source image, demands faithful reproduction of untouched parts', () => {
    const p = buildIllustrationRedesignPrompt('fix the arrow colors', '正文', { hasSourceImage: true });
    expect(p).toContain('attached as an image');
    expect(p).toContain('Reproduce its structure');
    expect(p).toContain('change only what the instruction asks for');
  });
});


// ============================================================================
// Image-model route (buildImagePrompt / illustrationRenderer)
// ============================================================================

describe('buildImagePrompt', () => {
  test('leads with the subject and carries the document context', () => {
    const p = buildImagePrompt('设计一张光合作用示意图', '本章讨论叶绿体与光反应。');
    expect(p).toContain('设计一张光合作用示意图');
    expect(p).toContain('本章讨论叶绿体与光反应。');
  });

  test('instructs the model to render no text', () => {
    // Every current image model garbles lettering; a labelled diagram full of
    // gibberish is worse than an unlabelled one.
    const p = buildImagePrompt('示意图', '正文');
    expect(p).toMatch(/Do not render any words, letters, numbers, labels, or captions/);
  });

  test('omits the SVG output contract wording entirely', () => {
    // An image model has no notion of an output contract: these rules would be
    // noise at best, and rendered as literal text in the picture at worst.
    const p = buildImagePrompt('设计并插入一张插图', '正文内容');
    expect(p).not.toMatch(/svg/i);
    expect(p).not.toMatch(/viewBox/i);
    expect(p).not.toContain('Output ONLY');
    expect(p).not.toMatch(/foreignObject/i);
    expect(p).not.toMatch(/markup/i);
    expect(p).not.toMatch(/code fences/i);
  });

  test('truncates an over-long context to keep the image prompt within API caps', () => {
    const longContext = 'x'.repeat(3000);
    const p = buildImagePrompt('示意图', longContext);

    expect(p).toContain('x'.repeat(1500));
    expect(p).not.toContain('x'.repeat(1501));
  });

  test('drops the context line when no document text is available', () => {
    const p = buildImagePrompt('设计一张插图', '');
    expect(p).not.toContain('Document topic for context');
    expect(p).toContain('设计一张插图');
  });

  test('drops the subject line when the instruction is empty, keeping a usable brief', () => {
    const p = buildImagePrompt('', '正文内容');
    expect(p).not.toContain('Subject:');
    expect(p).toContain('正文内容');
    expect(p).toMatch(/professional illustration/);
  });

  test('tolerates nullish arguments', () => {
    const p = buildImagePrompt(null, undefined);
    expect(typeof p).toBe('string');
    expect(p).toMatch(/professional illustration/);
    expect(p).not.toContain('null');
    expect(p).not.toContain('undefined');
  });
});

describe('illustrationRenderer', () => {
  test('explicit SVG/vector wording always picks the SVG route', () => {
    // Rule 1: an image model cannot deliver vector markup, so this wins even
    // when an image provider is configured and ready.
    for (const instruction of [
      '设计并增加SVG插图',
      'add an svg illustration',
      '画一张矢量图',
      '用矢量风格重画',
      '来一张向量图',
      'draw some line art',
      'a vector diagram please',
    ]) {
      expect(illustrationRenderer(instruction, true)).toBe('svg');
      expect(illustrationRenderer(instruction, false)).toBe('svg');
    }
  });

  test('a configured image model handles ordinary illustration requests', () => {
    expect(illustrationRenderer('设计示意图并插入', true)).toBe('image');
    expect(illustrationRenderer('给文章配一张插图', true)).toBe('image');
    expect(illustrationRenderer('insert an illustration at the top', true)).toBe('image');
  });

  test('without an image model everything falls back to SVG', () => {
    // Rule 3: existing installs keep working exactly as before.
    expect(illustrationRenderer('设计示意图并插入', false)).toBe('svg');
    expect(illustrationRenderer('给文章配一张插图', false)).toBe('svg');
  });

  test('a word merely containing the letters "svg" does not force the SVG route', () => {
    expect(illustrationRenderer('draw an svgish thing', true)).toBe('image');
  });

  test('nullish instructions follow the image-model readiness flag', () => {
    expect(illustrationRenderer(undefined, true)).toBe('image');
    expect(illustrationRenderer(null, false)).toBe('svg');
    expect(illustrationRenderer('', true)).toBe('image');
  });
});
