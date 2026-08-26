/**
 * Unit tests for the pure native-table creation protocol.
 */

const {
  TABLE_CREATION_LIMITS,
  TABLE_POSITIONS,
  TABLE_STYLES,
  DEFAULT_TABLE_CREATION_OPTIONS,
  inferTableCreationSpec,
  buildTableCreationPrompt,
  parseTableCreationResponse,
  validateTableCreationSpec,
} = require('../src/lib/table-ops.js');

function jsonSpec(overrides = {}) {
  return JSON.stringify({
    rows: [['Name', 'Value'], ['Alpha', '1']],
    position: 'end',
    headerRowCount: 1,
    style: 'tableGrid',
    autoFit: true,
    ...overrides,
  });
}

function errorCodes(result) {
  return result.errors.map((error) => error.code);
}

describe('table creation contract exports', () => {
  test('publishes stable allowlists, defaults, and conservative limits', () => {
    expect(TABLE_POSITIONS).toEqual(['start', 'end', 'before', 'after']);
    expect(TABLE_STYLES).toEqual(['tableGrid']);
    expect(DEFAULT_TABLE_CREATION_OPTIONS).toEqual({
      position: 'end',
      headerRowCount: 0,
      style: 'tableGrid',
      autoFit: true,
    });
    expect(TABLE_CREATION_LIMITS).toEqual({
      MAX_ROWS: 50,
      MAX_COLUMNS: 20,
      MAX_CELLS: 500,
      MAX_CELL_CHARS: 2000,
      MAX_TOTAL_CHARS: 50000,
      MAX_RESPONSE_CHARS: 100000,
    });
  });
});

describe('inferTableCreationSpec', () => {
  test.each([
    ['创建三行三列的空表格', 3, 3],
    ['插入3行3列空表', 3, 3],
    ['create a 3x3 table', 3, 3],
    ['Create 3 rows and 3 columns', 3, 3],
    ['Create 2 columns by 5 rows', 5, 2],
  ])('infers explicit dimensions from %s', (instruction, rowCount, columnCount) => {
    const spec = inferTableCreationSpec(instruction);

    expect(spec.rows).toHaveLength(rowCount);
    expect(spec.rows.every((row) => row.length === columnCount)).toBe(true);
    expect(spec.rows.flat().every((cell) => cell === '')).toBe(true);
  });

  test('uses the complete defaults for a plain empty-table request', () => {
    expect(inferTableCreationSpec('创建三行三列空表')).toEqual({
      rows: [
        ['', '', ''],
        ['', '', ''],
        ['', '', ''],
      ],
      position: 'end',
      headerRowCount: 0,
      style: 'tableGrid',
      autoFit: true,
    });
  });

  test.each([
    ['在文首创建2行2列表格', 'start'],
    ['开头插入2x2表格', 'start'],
    ['Create a 2x2 table at the start', 'start'],
    ['在文末创建2行2列表格', 'end'],
    ['末尾插入2x2表格', 'end'],
    ['Create a 2x2 table at the end', 'end'],
    ['Insert a 2x2 table before the selection', 'before'],
    ['在所选内容之前插入2行2列表格', 'before'],
    ['Insert a 2x2 table after the selection', 'after'],
    ['在所选内容之后插入2行2列表格', 'after'],
  ])('infers position from %s', (instruction, position) => {
    expect(inferTableCreationSpec(instruction).position).toBe(position);
  });

  test.each([
    '创建一个空表格',
    'make a table for the milestones',
    '创建三行的表格',
    'the document has 2026 entries',
    '',
  ])('returns null rather than guessing dimensions for %j', (instruction) => {
    expect(inferTableCreationSpec(instruction)).toBeNull();
  });

  test('returns null for explicit dimensions outside allocation limits', () => {
    expect(inferTableCreationSpec('创建51行2列表格')).toBeNull();
    expect(inferTableCreationSpec('create a 30x20 table')).toBeNull();
    expect(inferTableCreationSpec('create a 2x0 table')).toBeNull();
  });
});

describe('buildTableCreationPrompt', () => {
  test('embeds the request, optional context, schema, and strict output rules', () => {
    const prompt = buildTableCreationPrompt('生成季度数据表', 'Q1 revenue was 100.');

    expect(prompt).toContain('生成季度数据表');
    expect(prompt).toContain('Q1 revenue was 100.');
    expect(prompt).toContain('Output ONLY one JSON object');
    expect(prompt).toContain('"rows"');
    expect(prompt).toContain('"position": "start|end|before|after"');
    expect(prompt).toContain('"headerRowCount"');
    expect(prompt).toContain('"style": "tableGrid"');
    expect(prompt).toContain('"autoFit": true');
  });

  test('requires rectangular plain-text cells and forbids markup tables', () => {
    const prompt = buildTableCreationPrompt('create a comparison table');

    expect(prompt).toContain('rectangular matrix');
    expect(prompt).toContain('plain-text JSON string');
    expect(prompt).toContain('Do not use objects, nested arrays, null');
    expect(prompt).toContain('Markdown tables');
    expect(prompt).toContain('HTML tables/tags');
    expect(prompt).toContain(`${TABLE_CREATION_LIMITS.MAX_CELLS} cells`);
  });
});

describe('parseTableCreationResponse', () => {
  test('parses a complete valid spec', () => {
    expect(parseTableCreationResponse(jsonSpec())).toEqual({
      ok: true,
      spec: {
        rows: [['Name', 'Value'], ['Alpha', '1']],
        position: 'end',
        headerRowCount: 1,
        style: 'tableGrid',
        autoFit: true,
      },
      errors: [],
      warnings: [],
    });
  });

  test('tolerates a JSON fence and surrounding prose', () => {
    const raw = `Here is the table:\n\`\`\`json\n${jsonSpec({ position: 'before' })}\n\`\`\`\nDone.`;
    const result = parseTableCreationResponse(raw);

    expect(result.ok).toBe(true);
    expect(result.spec.position).toBe('before');
  });

  test.each([
    ['{broken json}', 'MALFORMED_JSON'],
    ['no JSON here', 'NO_JSON_OBJECT'],
    ['[["not", "an object"]]', 'RESPONSE_NOT_OBJECT'],
    ['null', 'RESPONSE_NOT_OBJECT'],
  ])('rejects malformed or non-object response %j', (raw, code) => {
    const result = parseTableCreationResponse(raw);
    expect(result.ok).toBe(false);
    expect(result.spec).toBeNull();
    expect(errorCodes(result)).toContain(code);
  });

  test.each([
    `[${jsonSpec()}]`,
    `Here is the table: [${jsonSpec()}] done.`,
  ])('rejects an object nested in a top-level JSON array', (raw) => {
    const result = parseTableCreationResponse(raw);
    expect(result.ok).toBe(false);
    expect(errorCodes(result)).toContain('RESPONSE_NOT_OBJECT');
  });

  test('applies defaults when optional fields are absent', () => {
    const result = parseTableCreationResponse('{"rows":[["a","b"]]}');
    expect(result).toEqual({
      ok: true,
      spec: {
        rows: [['a', 'b']],
        position: 'end',
        headerRowCount: 0,
        style: 'tableGrid',
        autoFit: true,
      },
      errors: [],
      warnings: [],
    });
  });

  test('coerces finite number and boolean cell primitives with warnings', () => {
    const result = parseTableCreationResponse('{"rows":[[1,true,false,2.5]]}');

    expect(result.ok).toBe(true);
    expect(result.spec.rows).toEqual([['1', 'true', 'false', '2.5']]);
    expect(result.warnings).toEqual([
      expect.objectContaining({ code: 'CELL_PRIMITIVE_COERCED', path: 'rows[0][0]' }),
      expect.objectContaining({ code: 'CELL_PRIMITIVE_COERCED', path: 'rows[0][1]' }),
      expect.objectContaining({ code: 'CELL_PRIMITIVE_COERCED', path: 'rows[0][2]' }),
      expect.objectContaining({ code: 'CELL_PRIMITIVE_COERCED', path: 'rows[0][3]' }),
    ]);
  });

  test('rejects ragged rows without padding or truncating', () => {
    const result = parseTableCreationResponse('{"rows":[["a","b"],["c"]]}');

    expect(result.ok).toBe(false);
    expect(result.spec).toBeNull();
    expect(result.errors).toContainEqual(expect.objectContaining({
      code: 'RAGGED_ROWS',
      path: 'rows[1]',
    }));
  });

  test.each([
    ['{"rows":[[{}]]}', 'rows[0][0]'],
    ['{"rows":[[["nested"]]]}', 'rows[0][0]'],
    ['{"rows":[[null]]}', 'rows[0][0]'],
  ])('rejects object, nested-array, and null cells', (raw, path) => {
    const result = parseTableCreationResponse(raw);
    expect(result.ok).toBe(false);
    expect(result.errors).toContainEqual(expect.objectContaining({ code: 'CELL_NOT_TEXT', path }));
  });

  test.each([
    [{ position: 'middle' }, 'UNSUPPORTED_POSITION'],
    [{ style: 'LightShading' }, 'UNSUPPORTED_STYLE'],
    [{ autoFit: 'true' }, 'AUTO_FIT_NOT_BOOLEAN'],
    [{ headerRowCount: 1.5 }, 'INVALID_HEADER_ROW_COUNT'],
    [{ headerRowCount: 3 }, 'HEADER_ROW_COUNT_OUT_OF_RANGE'],
  ])('rejects unsupported or mistyped options: %j', (override, code) => {
    const result = parseTableCreationResponse(jsonSpec(override));
    expect(result.ok).toBe(false);
    expect(errorCodes(result)).toContain(code);
  });

  test.each([
    [{ rows: [] }, 'ROW_COUNT_OUT_OF_RANGE'],
    [{ rows: [[]] }, 'COLUMN_COUNT_OUT_OF_RANGE'],
    [{ rows: Array.from({ length: 51 }, () => ['']) }, 'ROW_COUNT_OUT_OF_RANGE'],
    [{ rows: [Array(21).fill('')] }, 'COLUMN_COUNT_OUT_OF_RANGE'],
    [{ rows: Array.from({ length: 26 }, () => Array(20).fill('')) }, 'CELL_COUNT_OUT_OF_RANGE'],
  ])('rejects zero and oversized dimensions: %s', (override, code) => {
    const result = parseTableCreationResponse(jsonSpec(override));
    expect(result.ok).toBe(false);
    expect(errorCodes(result)).toContain(code);
  });

  test('rejects excessive per-cell and total text', () => {
    const longCell = parseTableCreationResponse(jsonSpec({
      rows: [['x'.repeat(TABLE_CREATION_LIMITS.MAX_CELL_CHARS + 1)]],
    }));
    expect(errorCodes(longCell)).toContain('CELL_TEXT_TOO_LARGE');

    const totalText = parseTableCreationResponse(jsonSpec({
      rows: Array.from({ length: 25 }, () => Array(20).fill('x'.repeat(101))),
    }));
    expect(errorCodes(totalText)).toContain('TOTAL_TEXT_TOO_LARGE');
  });

  test('rejects responses over the raw response cap before parsing', () => {
    const result = parseTableCreationResponse('x'.repeat(TABLE_CREATION_LIMITS.MAX_RESPONSE_CHARS + 1));
    expect(errorCodes(result)).toEqual(['RESPONSE_TOO_LARGE']);
  });

  test('reports ignored unknown properties as structured warnings', () => {
    const result = parseTableCreationResponse(jsonSpec({ explanation: 'extra' }));
    expect(result.ok).toBe(true);
    expect(result.spec).not.toHaveProperty('explanation');
    expect(result.warnings).toContainEqual(expect.objectContaining({
      code: 'UNKNOWN_PROPERTY',
      path: 'explanation',
    }));
  });
});

describe('validateTableCreationSpec', () => {
  test('rejects a non-object candidate with a structured error', () => {
    expect(validateTableCreationSpec([['a']])).toEqual({
      ok: false,
      spec: null,
      errors: [{
        code: 'SPEC_NOT_OBJECT',
        path: '$',
        message: 'Table creation spec must be a JSON object',
      }],
      warnings: [],
    });
  });
});
