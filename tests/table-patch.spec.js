/**
 * Unit tests for src/lib/table-patch.js — the coordinate-addressed table
 * revision protocol (prompt building, response parsing/validation, row-op
 * ordering).
 */

const {
  formatCellGrid,
  buildTableUserPrompt,
  parseTablePatchResponse,
  planRowOpOrder,
} = require('../src/lib/table-patch.js');

const DIMS_3X2 = { rowCount: 3, colCount: 2 };
const ORIGINALS_3X2 = [
  ['Header A', 'Header B'],
  ['old a', 'b'],
  ['c', 'd'],
];

describe('formatCellGrid', () => {
  test('lists one coordinate line per cell', () => {
    const grid = formatCellGrid([
      { row: 1, col: 1, text: 'Header A' },
      { row: 1, col: 2, text: 'Header B' },
      { row: 2, col: 1, text: 'old a' },
    ]);
    expect(grid).toBe('[R1C1] Header A\n[R1C2] Header B\n[R2C1] old a');
  });

  test('marks merge-covered slots as read-only', () => {
    const grid = formatCellGrid([
      { row: 1, col: 1, text: 'Wide header' },
      { row: 1, col: 2, text: '', merged: true },
    ]);
    expect(grid).toBe('[R1C1] Wide header\n[R1C2] (merged — read-only) ');
  });
});

describe('buildTableUserPrompt', () => {
  test('embeds instruction, grid and JSON rules', () => {
    const cells = [{ row: 2, col: 1, text: 'old a' }];
    const prompt = buildTableUserPrompt('Fix the numbers', cells, DIMS_3X2);

    expect(prompt).toContain('Fix the numbers');
    expect(prompt).toContain('[R2C1] old a');
    expect(prompt).toContain('"cells"');
    expect(prompt).toContain('"rowOps"');
    expect(prompt).toContain('insertAfter');
    expect(prompt).toContain('exactly 2 entries');
    expect(prompt).toContain('3 rows x 2 columns');
    expect(prompt).not.toContain('MERGED CELLS');
  });

  test('adds merged-cell rules when any covered slot is merge-covered', () => {
    const cells = [
      { row: 1, col: 1, text: 'Wide header' },
      { row: 1, col: 2, text: '', merged: true },
    ];
    const prompt = buildTableUserPrompt('Fix the header', cells, DIMS_3X2);

    expect(prompt).toContain('(merged — read-only)');
    expect(prompt).toContain('MERGED CELLS');
    expect(prompt).toContain('"rowOps" MUST be an empty array');
  });
});

describe('parseTablePatchResponse', () => {
  test('parses a clean patch and sorts cells in document order', () => {
    const raw = '{"cells":[{"row":3,"col":1,"text":"c2"},{"row":2,"col":1,"text":"new a"}]}';
    const { cells, rowOps, warnings } = parseTablePatchResponse(raw, { ...DIMS_3X2, originals: ORIGINALS_3X2 });

    expect(cells).toEqual([
      { row: 2, col: 1, text: 'new a' },
      { row: 3, col: 1, text: 'c2' },
    ]);
    expect(rowOps).toEqual([]);
    expect(warnings).toEqual([]);
  });

  test('drops patch entries targeting merge-covered (shadow) coordinates', () => {
    const raw = '{"cells":[{"row":1,"col":2,"text":"x"},{"row":2,"col":1,"text":"new a"}]}';
    const { cells, warnings } = parseTablePatchResponse(raw, {
      ...DIMS_3X2,
      originals: ORIGINALS_3X2,
      shadowCoords: new Set(['1,2']),
    });

    expect(cells).toEqual([{ row: 2, col: 1, text: 'new a' }]);
    expect(warnings).toEqual([
      'Cell R1C2 is covered by a merged cell — not editable, dropped',
    ]);
  });

  test('tolerates markdown fences, surrounding prose and trailing commas', () => {
    const raw = 'Sure! Here is the patch:\n```json\n{"cells":[{"row":2,"col":1,"text":"new a",},],}\n```\nDone.';
    const { cells } = parseTablePatchResponse(raw, { ...DIMS_3X2, originals: ORIGINALS_3X2 });
    expect(cells).toEqual([{ row: 2, col: 1, text: 'new a' }]);
  });

  test('throws when no JSON object is present', () => {
    expect(() => parseTablePatchResponse('no json here', DIMS_3X2)).toThrow(/no JSON object/);
    expect(() => parseTablePatchResponse('{broken json}', DIMS_3X2)).toThrow(/parse failed/);
  });

  test('drops out-of-bounds coordinates with warnings', () => {
    const raw = '{"cells":[{"row":99,"col":1,"text":"x"},{"row":2,"col":0,"text":"y"},{"row":2,"col":1,"text":"new a"}]}';
    const { cells, warnings } = parseTablePatchResponse(raw, { ...DIMS_3X2, originals: ORIGINALS_3X2 });

    expect(cells).toEqual([{ row: 2, col: 1, text: 'new a' }]);
    expect(warnings).toHaveLength(2);
    expect(warnings[0]).toMatch(/out of bounds/);
  });

  test('drops no-op cells (text matches the original, ignoring whitespace)', () => {
    const raw = '{"cells":[{"row":2,"col":1,"text":"old a"},{"row":2,"col":2,"text":"  b  "},{"row":3,"col":1,"text":"c2"}]}';
    const { cells } = parseTablePatchResponse(raw, { ...DIMS_3X2, originals: ORIGINALS_3X2 });
    expect(cells).toEqual([{ row: 3, col: 1, text: 'c2' }]);
  });

  test('keeps no-op-looking cells when originals are not provided', () => {
    const raw = '{"cells":[{"row":2,"col":1,"text":"old a"}]}';
    const { cells } = parseTablePatchResponse(raw, DIMS_3X2);
    expect(cells).toEqual([{ row: 2, col: 1, text: 'old a' }]);
  });

  test('conflicting duplicate cell entries are all dropped (never last-wins)', () => {
    const raw = '{"cells":[{"row":2,"col":1,"text":"first"},{"row":2,"col":1,"text":"second"}]}';
    const { cells, warnings } = parseTablePatchResponse(raw, { ...DIMS_3X2, originals: ORIGINALS_3X2 });
    expect(cells).toEqual([]);
    expect(warnings[0]).toMatch(/Conflicting duplicate/);
  });

  test('identical duplicate cell entries are coalesced with a warning', () => {
    const raw = '{"cells":[{"row":2,"col":1,"text":"new a"},{"row":2,"col":1,"text":"new a"}]}';
    const { cells, warnings } = parseTablePatchResponse(raw, { ...DIMS_3X2, originals: ORIGINALS_3X2 });
    expect(cells).toEqual([{ row: 2, col: 1, text: 'new a' }]);
    expect(warnings[0]).toMatch(/coalesced/);
  });

  test('coerces numeric cell values to text', () => {
    const raw = '{"cells":[{"row":2,"col":1,"text":1.5}]}';
    const { cells } = parseTablePatchResponse(raw, { ...DIMS_3X2 });
    expect(cells).toEqual([{ row: 2, col: 1, text: '1.5' }]);
  });

  test('parses row ops and pre-sorts them descending', () => {
    const raw = '{"rowOps":[{"op":"insertAfter","row":1,"values":["n1","n2"]},{"op":"delete","row":3}]}';
    const { rowOps, warnings } = parseTablePatchResponse(raw, DIMS_3X2);
    expect(rowOps).toEqual([
      { op: 'delete', row: 3 },
      { op: 'insertAfter', row: 1, values: ['n1', 'n2'] },
    ]);
    expect(warnings).toEqual([]);
  });

  test('validates row ops: unknown op, bad row, missing values', () => {
    const raw = '{"rowOps":[{"op":"merge","row":1},{"op":"delete","row":42},{"op":"insertAfter","row":2}]}';
    const { rowOps, warnings } = parseTablePatchResponse(raw, DIMS_3X2);
    expect(rowOps).toEqual([]);
    expect(warnings).toHaveLength(3);
  });

  test('insert values must match the column count exactly (no silent pad/truncate)', () => {
    const raw = '{"rowOps":[{"op":"insertBefore","row":1,"values":["only"]},{"op":"insertAfter","row":2,"values":["a","b","c"]},{"op":"insertAfter","row":3,"values":["x","y"]}]}';
    const { rowOps, warnings } = parseTablePatchResponse(raw, DIMS_3X2);
    expect(rowOps).toEqual([
      { op: 'insertAfter', row: 3, values: ['x', 'y'] },
    ]);
    expect(warnings).toHaveLength(2);
    expect(warnings[0]).toMatch(/exactly 2 required/);
  });

  test('non-array cells/rowOps are ignored with warnings', () => {
    const { cells, rowOps, warnings } = parseTablePatchResponse('{"cells":{},"rowOps":{}}', DIMS_3X2);
    expect(cells).toEqual([]);
    expect(rowOps).toEqual([]);
    expect(warnings).toHaveLength(2);
  });

  test('allowedBounds restrict cells to the selected region', () => {
    const raw = '{"cells":[{"row":1,"col":1,"text":"outside"},{"row":2,"col":2,"text":"inside"}]}';
    const { cells, warnings } = parseTablePatchResponse(raw, {
      ...DIMS_3X2,
      allowedBounds: { startRow: 2, endRow: 3, startCol: 2, endCol: 2 },
    });
    expect(cells).toEqual([{ row: 2, col: 2, text: 'inside' }]);
    expect(warnings[0]).toMatch(/outside allowedBounds/);
  });

  test('invalid allowedBounds reject the entire patch', () => {
    const raw = '{"cells":[{"row":1,"col":1,"text":"x"}]}';
    const { cells, warnings } = parseTablePatchResponse(raw, {
      ...DIMS_3X2,
      allowedBounds: { startRow: 3, endRow: 1, startCol: 1, endCol: 2 },
    });
    expect(cells).toEqual([]);
    expect(warnings[0]).toMatch(/allowedBounds/);
  });

  test('row ops are rejected for partial-width or disallowed selections', () => {
    const raw = '{"rowOps":[{"op":"delete","row":2}]}';
    const partial = parseTablePatchResponse(raw, {
      ...DIMS_3X2,
      allowedBounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 1, allowRowOps: true },
    });
    expect(partial.rowOps).toEqual([]);
    expect(partial.warnings[0]).toMatch(/full-width/);

    const disallowed = parseTablePatchResponse(raw, {
      ...DIMS_3X2,
      allowedBounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
    });
    expect(disallowed.rowOps).toEqual([]);
    expect(disallowed.warnings[0]).toMatch(/not allowed/);
  });

  test('full-width allowedBounds with allowRowOps accept in-range row ops', () => {
    const raw = '{"rowOps":[{"op":"delete","row":2},{"op":"insertAfter","row":1,"values":["a","b"]}]}';
    const { rowOps, warnings } = parseTablePatchResponse(raw, {
      ...DIMS_3X2,
      allowedBounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2, allowRowOps: true },
    });
    expect(rowOps).toEqual([
      { op: 'delete', row: 2 },
      { op: 'insertAfter', row: 1, values: ['a', 'b'] },
    ]);
    expect(warnings).toEqual([]);
  });

  test('conflicting duplicate row ops are all dropped', () => {
    const raw = '{"rowOps":[{"op":"insertAfter","row":1,"values":["a","b"]},{"op":"insertAfter","row":1,"values":["x","y"]}]}';
    const { rowOps, warnings } = parseTablePatchResponse(raw, DIMS_3X2);
    expect(rowOps).toEqual([]);
    expect(warnings[0]).toMatch(/Conflicting duplicate row op/);
  });

  test('deleting every existing row is rejected as a whole', () => {
    const raw = '{"rowOps":[{"op":"delete","row":1},{"op":"delete","row":2},{"op":"delete","row":3}]}';
    const { rowOps, warnings } = parseTablePatchResponse(raw, DIMS_3X2);
    expect(rowOps).toEqual([]);
    expect(warnings[0]).toMatch(/delete all existing rows/);
  });

  test('unsupported colOps are warned and ignored', () => {
    const raw = '{"cells":[{"row":2,"col":1,"text":"new a"}],"colOps":[{"op":"delete","col":1}]}';
    const { cells, warnings } = parseTablePatchResponse(raw, DIMS_3X2);
    expect(cells).toHaveLength(1);
    expect(warnings[0]).toMatch(/colOps/);
  });

  test('oversized responses are rejected before parsing', () => {
    const { cells, rowOps, warnings } = parseTablePatchResponse('x'.repeat(300 * 1024), DIMS_3X2);
    expect(cells).toEqual([]);
    expect(rowOps).toEqual([]);
    expect(warnings[0]).toMatch(/character limit/);
  });

  test('entry-count caps drop the whole bucket with a warning', () => {
    const manyCells = Array.from({ length: 600 }, (_, i) => ({ row: 1, col: 1, text: `v${i}` }));
    const { cells, warnings } = parseTablePatchResponse(JSON.stringify({ cells: manyCells }), DIMS_3X2);
    expect(cells).toEqual([]);
    expect(warnings[0]).toMatch(/limit is 512/);
  });
});

describe('planRowOpOrder', () => {
  test('sorts by descending row so earlier ops never shift later coordinates', () => {
    const ordered = planRowOpOrder([
      { op: 'insertAfter', row: 1 },
      { op: 'delete', row: 5 },
      { op: 'delete', row: 2 },
    ]);
    expect(ordered.map((o) => `${o.op}:${o.row}`)).toEqual(['delete:5', 'delete:2', 'insertAfter:1']);
  });

  test('same-row ties resolve inserts before deletes (replace-row semantics)', () => {
    const ordered = planRowOpOrder([
      { op: 'delete', row: 3 },
      { op: 'insertAfter', row: 3 },
    ]);
    expect(ordered.map((o) => o.op)).toEqual(['insertAfter', 'delete']);
  });

  test('does not mutate the input array', () => {
    const input = [{ op: 'delete', row: 1 }, { op: 'delete', row: 2 }];
    planRowOpOrder(input);
    expect(input.map((o) => o.row)).toEqual([1, 2]);
  });
});
