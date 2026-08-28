/**
 * Specs for src/lib/table-model.js — the table tool-loop draft model.
 * Covers op validation (bounds, shadows, merged rules, no-ops) and the
 * translation into the existing tablePatch shape (apply order preserved).
 */

const { createTableModel, executeTableTool, TABLE_TOOL_SPECS } = require('../src/lib/table-model.js');

const REGION = {
    rowCount: 3,
    colCount: 2,
    values: [
        ['Header A', 'Header B'],
        ['old a', 'b'],
        ['c', 'd'],
    ],
    bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 2 },
    merged: false,
    shadowKeys: new Set(),
};

describe('createTableModel', () => {
    test('get_state lists the covered grid and row-op permission', () => {
        const model = createTableModel(REGION);
        const { ok, result } = model.getState();
        expect(ok).toBe(true);
        expect(result.rowOpsAllowed).toBe(true);
        expect(result.grid).toContain('[R2C1] old a');
        expect(result.grid).toContain('[R1C1] Header A');
        expect(result.pendingOps).toEqual([]);
    });

    test('get_state marks merged shadow slots read-only', () => {
        const model = createTableModel({
            ...REGION,
            merged: true,
            shadowKeys: new Set(['1,2']),
        });
        const { result } = model.getState();
        expect(result.grid).toContain('[R1C2] (merged — read-only)');
        expect(result.mergedTable).toBe(true);
        expect(result.rowOpsAllowed).toBe(false);
    });

    test('set_cell records an edit and later edits overwrite the same cell', () => {
        const model = createTableModel(REGION);
        expect(model.setCell(2, 1, 'first')).toEqual({ ok: true, result: { set: 'R2C1' } });
        expect(model.setCell(2, 1, 'second')).toEqual({ ok: true, result: { set: 'R2C1' } });
        const patch = model.toTablePatch();
        expect(patch.cells).toEqual([{ row: 2, col: 1, text: 'second' }]);
    });

    test('set_cell rejects out-of-bounds, shadow, and identical text', () => {
        const model = createTableModel({
            ...REGION,
            bounds: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 },
        });
        expect(model.setCell(1, 1, 'x').error).toMatch(/outside the covered region/);
        expect(model.setCell(99, 1, 'x').error).toMatch(/out of bounds/);
        expect(model.setCell(2, 1, 42).error).toMatch(/must be a string/);

        const identical = model.setCell(2, 1, 'old a');
        expect(identical.ok).toBe(true);
        expect(identical.result.note).toMatch(/already holds that text/);
        expect(model.opCount).toBe(0);

        const merged = createTableModel({ ...REGION, merged: true, shadowKeys: new Set(['2,1']) });
        expect(merged.setCell(2, 1, 'x').error).toMatch(/merged cell/);
    });

    test('set_cell rejects rows with a pending delete', () => {
        const model = createTableModel(REGION);
        expect(model.deleteRow(2).ok).toBe(true);
        expect(model.setCell(2, 1, 'x').error).toMatch(/pending delete/);
        expect(model.setCell(3, 1, 'e').ok).toBe(true);
    });

    test('insert_row validates position, anchor, and values length', () => {
        const model = createTableModel(REGION);
        expect(model.insertRow({ position: 'beside', row: 1, values: ['a', 'b'] }).error).toMatch(/after" or "before/);
        expect(model.insertRow({ position: 'after', row: 0, values: ['a', 'b'] }).error).toMatch(/out of bounds/);
        expect(model.insertRow({ position: 'after', row: 1, values: ['only one'] }).error).toMatch(/exactly 2/);
        expect(model.insertRow({ position: 'after', row: 1, values: ['a', 5] }).error).toMatch(/must be a string/);
        expect(model.insertRow({ position: 'after', row: 1, values: ['a', 'b'] }).ok).toBe(true);
    });

    test('row ops are rejected on merged tables and partial-width selections', () => {
        const merged = createTableModel({ ...REGION, merged: true, shadowKeys: new Set() });
        expect(merged.deleteRow(1).error).toMatch(/merged cells/);
        expect(merged.insertRow({ position: 'after', row: 1, values: ['a', 'b'] }).error).toMatch(/merged cells/);

        const narrow = createTableModel({
            ...REGION,
            bounds: { startRow: 1, endRow: 3, startCol: 1, endCol: 1 },
        });
        expect(narrow.deleteRow(1).error).toMatch(/full-width/);
    });

    test('delete_row refuses to remove every row and coalesces duplicates', () => {
        const model = createTableModel({ ...REGION, values: [['a', 'b']] , rowCount: 1 });
        expect(model.deleteRow(1).error).toMatch(/every row/);

        const model2 = createTableModel(REGION);
        expect(model2.deleteRow(3).ok).toBe(true);
        const again = model2.deleteRow(3);
        expect(again.ok).toBe(true);
        expect(again.result.note).toMatch(/already has a pending delete/);
        expect(model2.toTablePatch().rowOps).toEqual([{ op: 'delete', row: 3 }]);
    });

    test('toTablePatch translates edits + row ops into apply order', () => {
        const model = createTableModel(REGION);
        model.setCell(2, 1, 'new a');
        model.setCell(3, 2, 'd2'); // recorded before the delete below
        model.insertRow({ position: 'after', row: 1, values: ['n1', 'n2'] });
        model.deleteRow(3);
        // Post-delete edits on row 3 are rejected outright.
        expect(model.setCell(3, 2, 'later').error).toMatch(/pending delete/);

        const patch = model.toTablePatch();
        // Cells in a deleted row are dropped from the patch.
        expect(patch.cells).toEqual([{ row: 2, col: 1, text: 'new a' }]);
        // Row ops in descending application order (delete 3 before insert@1).
        expect(patch.rowOps).toEqual([
            { op: 'delete', row: 3 },
            { op: 'insertAfter', row: 1, values: ['n1', 'n2'] },
        ]);
        expect(patch.bounds).toEqual(REGION.bounds);
        expect(patch.originals).toEqual(REGION.values);
    });
});

describe('executeTableTool', () => {
    test('dispatches tool names to model methods', () => {
        const model = createTableModel(REGION);
        expect(executeTableTool(model, 'get_state', {}).ok).toBe(true);
        expect(executeTableTool(model, 'set_cell', { row: 2, col: 1, text: 'x' }).ok).toBe(true);
        expect(executeTableTool(model, 'insert_row', { position: 'after', row: 1, values: ['a', 'b'] }).ok).toBe(true);
        expect(executeTableTool(model, 'delete_row', { row: 3 }).ok).toBe(true);
        expect(executeTableTool(model, 'bogus', {}).ok).toBe(false);
    });

    test('TABLE_TOOL_SPECS are registry-valid', () => {
        expect(TABLE_TOOL_SPECS.map((t) => t.name)).toEqual([
            'get_state', 'set_cell', 'insert_row', 'delete_row', 'merge_cells',
            'set_table_style', 'set_borders', 'set_cell_format', 'set_font',
            'set_header_row', 'set_layout', 'set_column_widths',
        ]);
    });

    test('merge_cells stages a rectangular merge and marks away-cells read-only', () => {
        const model = createTableModel(REGION);
        expect(model.mergeCells({ row: 2, col: 1, rows: 2, cols: 2 }))
            .toEqual({ ok: true, result: { merged: 'R2C1–R3C2 (anchor R2C1)' } });

        // Away cells (non-anchor) become read-only — set_cell rejects them.
        expect(model.setCell(2, 2, 'x').error).toMatch(/merged away/);
        expect(model.setCell(3, 1, 'x').error).toMatch(/merged away/);
        // The anchor stays editable.
        expect(model.setCell(2, 1, 'merged cell').ok).toBe(true);

        // Grid reflects the merge.
        const { result } = model.getState();
        expect(result.grid).toContain('[R2C2] (merged — read-only)');
        expect(result.pendingOps).toEqual([
            { tool: 'set_cell', row: 2, col: 1, text: 'merged cell' },
            { tool: 'merge_cells', op: 'merge', startRow: 2, startCol: 1, endRow: 3, endCol: 2 },
        ]);

        const patch = model.toTablePatch();
        expect(patch.merges).toEqual([{ op: 'merge', startRow: 2, startCol: 1, endRow: 3, endCol: 2 }]);
    });

    test('merge_cells validation: bounds, single-cell, second merge, existing merge', () => {
        const model = createTableModel(REGION);
        expect(model.mergeCells({ row: 1, col: 1, rows: 1, cols: 2 }).ok).toBe(true);
        expect(model.mergeCells({ row: 1, col: 1, rows: 2, cols: 2 }).error).toMatch(/Only one cell merge/);
        expect(model.mergeCells({ row: 1, col: 1, rows: 1, cols: 1 }).error).toMatch(/single cell/);
        expect(model.mergeCells({ row: 9, col: 1, rows: 2, cols: 2 }).error).toMatch(/outside the 3x2 table/);

        const merged = createTableModel({ ...REGION, merged: true, shadowKeys: new Set(['1,2']) });
        expect(merged.mergeCells({ row: 1, col: 1, rows: 2, cols: 2 }).error).toMatch(/existing merged cell/);
    });

    test('merge_cells rejects a region overlapping a pending-delete row', () => {
        const model = createTableModel(REGION);
        expect(model.deleteRow(3).ok).toBe(true);
        expect(model.mergeCells({ row: 2, col: 1, rows: 2, cols: 2 }).error).toMatch(/pending delete/);
    });

    test('a staged merge disables further row operations', () => {
        const model = createTableModel(REGION);
        expect(model.mergeCells({ row: 1, col: 1, rows: 1, cols: 2 }).ok).toBe(true);
        expect(model.insertRow({ position: 'after', row: 2, values: ['a', 'b'] }).error)
            .toMatch(/cell merge is pending/);
        expect(model.deleteRow(2).error).toMatch(/cell merge is pending/);
    });

    test('executeTableTool dispatches merge_cells', () => {
        const model = createTableModel(REGION);
        const out = executeTableTool(model, 'merge_cells', { row: 1, col: 1, rows: 2, cols: 1 });
        expect(out.ok).toBe(true);
        expect(model.opCount).toBe(1);
    });
});

describe('table style tools', () => {
    test('get_state reports the style snapshot and pending style ops', () => {
        const model = createTableModel({
            ...REGION,
            style: {
                styleBuiltIn: 'TableGrid', headerRowCount: 1, alignment: 'Left',
                horizontalAlignment: 'Left', verticalAlignment: 'Top', shadingColor: null,
                font: { name: 'Calibri', size: 11, bold: false },
                borders: { top: { type: 'single', width: 0.5 }, insideH: { type: 'none' } },
            },
        });
        expect(model.setTableStyle({ style: 'grid table 4 accent 1', bandedRows: true }).ok).toBe(true);
        const { result } = model.getState();
        expect(result.style).toContain('TableGrid');
        expect(result.style).toContain('header rows: 1');
        expect(result.style).toContain('top=single 0.5pt');
        expect(result.style).toContain('insideH=none');
        expect(result.pendingOps).toContainEqual({
            type: 'tableStyle', tool: 'set_table_style',
            style: 'GridTable4_Accent1', bandedRows: true,
        });

        // No snapshot → honest "unavailable" line.
        const bare = createTableModel(REGION).getState().result;
        expect(bare.style).toMatch(/snapshot unavailable/);
    });

    test('set_table_style canonicalizes names and rejects unknown ones', () => {
        const model = createTableModel(REGION);
        expect(model.setTableStyle({ style: 'plain table 3' }).result.staged).toContain('PlainTable3');
        expect(model.setTableStyle({ bandedRows: false }).ok).toBe(true);
        expect(model.setTableStyle({ style: 'FancyTable' }).error).toMatch(/Unknown built-in table style/);
        expect(model.setTableStyle({ bandedRows: 'yes' }).error).toMatch(/true or false/);
        expect(model.setTableStyle({}).error).toMatch(/at least one/);
    });

    test('set_borders expands aliases and validates row targeting', () => {
        const model = createTableModel(REGION);
        const out = model.setBorders({ borders: { top: { type: 'single', width: 1.5 }, inside: 'none' } });
        expect(out.ok).toBe(true);
        const patch = model.toTablePatch();
        expect(patch.styleOps[0].borders).toMatchObject({
            top: { type: 'single', width: 1.5 },
            insideH: { type: 'none' },
            insideV: { type: 'none' },
        });

        expect(model.setBorders({ row: 1, borders: { bottom: 'single' } }).result.staged)
            .toContain('row 1');
        expect(model.setBorders({ row: 9, borders: { top: 'single' } }).error).toMatch(/out of bounds/);
        expect(model.setBorders({ borders: {} }).error).toMatch(/no border locations/);

        const deleted = createTableModel(REGION);
        deleted.deleteRow(2);
        expect(deleted.setBorders({ row: 2, borders: { top: 'single' } }).error).toMatch(/pending delete/);
    });

    test('set_cell_format targets cells, bands, blocks, and the whole table', () => {
        const model = createTableModel(REGION);
        expect(model.setCellFormat({ row: 1, shadingColor: 'light blue' }).ok).toBe(true);
        expect(model.setCellFormat({ row: 2, horizontalAlignment: 'middle' }).ok).toBe(true);
        expect(model.setCellFormat({ col: 1, verticalAlignment: 'bottom' }).ok).toBe(true);
        expect(model.setCellFormat({ row: 2, col: 2, rows: 2, cols: 1, shadingColor: 'auto' }).ok).toBe(true);
        expect(model.setCellFormat({ shadingColor: '#FFFFFF', horizontalAlignment: 'right' }).ok).toBe(true);

        const styleOps = model.toTablePatch().styleOps;
        expect(styleOps[0]).toMatchObject({ type: 'cellFormat', region: { startRow: 1, endRow: 1, startCol: 1, endCol: 2 }, shadingColor: '#ADD8E6' });
        expect(styleOps[1].horizontalAlignment).toBe('centered');
        expect(styleOps[4].region).toBeNull();

        expect(model.setCellFormat({ row: 1, shadingColor: 'zebra' }).error).toMatch(/Invalid shadingColor/);
        expect(model.setCellFormat({ row: 1, horizontalAlignment: 'diagonal' }).error).toMatch(/horizontalAlignment/);
        expect(model.setCellFormat({ row: 1 }).error).toMatch(/at least one/);
    });

    test('set_cell_format / set_font reject pending-delete rows', () => {
        const model = createTableModel(REGION);
        model.deleteRow(2);
        expect(model.setCellFormat({ row: 2, shadingColor: 'red' }).error).toMatch(/pending delete/);
        expect(model.setFont({ row: 2, font: { bold: true } }).error).toMatch(/pending delete/);
        // Whole-table stays allowed (it maps to table-level properties).
        expect(model.setCellFormat({ shadingColor: 'red' }).ok).toBe(true);
    });

    test('set_font validates payload and whole-table targeting', () => {
        const model = createTableModel(REGION);
        expect(model.setFont({ row: 1, font: { bold: true, size: 11 } }).ok).toBe(true);
        expect(model.setFont({ font: { name: 'SimSun', color: 'black' } }).ok).toBe(true);
        const styleOps = model.toTablePatch().styleOps;
        expect(styleOps[0]).toMatchObject({ type: 'font', font: { bold: true, size: 11 } });
        expect(styleOps[1].region).toBeNull();

        expect(model.setFont({ row: 1, font: { bold: 'yes' } }).error).toMatch(/true or false/);
        expect(model.setFont({ row: 1, font: {} }).error).toMatch(/no supported keys/);
        expect(model.setFont({ row: 9, font: { bold: true } }).error).toMatch(/outside the covered region/);
    });

    test('set_header_row requires the region to include row 1 and validates rows', () => {
        const model = createTableModel(REGION);
        const out = model.setHeaderRow({ font: { bold: true }, shadingColor: '#DEEBF7' });
        expect(out.ok).toBe(true);
        expect(model.toTablePatch().styleOps[0]).toEqual({
            type: 'headerRow', tool: 'set_header_row', rows: 1,
            font: { bold: true }, shadingColor: '#DEEBF7',
        });
        expect(model.setHeaderRow({ rows: 4 }).error).toMatch(/1–3/);
        expect(model.setHeaderRow({ shadingColor: 'mud' }).error).toMatch(/Invalid shadingColor/);

        const partial = createTableModel({ ...REGION, bounds: { startRow: 2, endRow: 3, startCol: 1, endCol: 2 } });
        expect(partial.setHeaderRow({}).error).toMatch(/must include row 1/);

        const deleted = createTableModel(REGION);
        deleted.deleteRow(1);
        expect(deleted.setHeaderRow({}).error).toMatch(/pending delete/);
    });

    test('set_layout validates each field', () => {
        const model = createTableModel(REGION);
        expect(model.setLayout({ alignment: 'center', autoFitWindow: true }).ok).toBe(true);
        expect(model.setLayout({ cellPaddingPt: 5.25, widthPt: 450 }).ok).toBe(true);
        expect(model.toTablePatch().styleOps[0]).toMatchObject({ alignment: 'centered', autoFitWindow: true });

        expect(model.setLayout({ alignment: 'justify' }).error).toMatch(/alignment/);
        expect(model.setLayout({ widthPt: 3 }).error).toMatch(/widthPt/);
        expect(model.setLayout({ cellPaddingPt: 500 }).error).toMatch(/cellPaddingPt/);
        expect(model.setLayout({ autoFitWindow: 'yes' }).error).toMatch(/true or false/);
        expect(model.setLayout({}).error).toMatch(/at least one/);
    });

    test('set_column_widths requires uniform tables and exact lengths', () => {
        const model = createTableModel(REGION);
        expect(model.setColumnWidths({ widthsPt: [120, 80] }).ok).toBe(true);
        expect(model.toTablePatch().styleOps[0].widthsPt).toEqual([120, 80]);

        expect(model.setColumnWidths({ widthsPt: [100] }).error).toMatch(/exactly 2/);
        expect(model.setColumnWidths({ widthsPt: [100, 2] }).error).toMatch(/Every width/);

        const merged = createTableModel({ ...REGION, merged: true, shadowKeys: new Set() });
        expect(merged.setColumnWidths({ widthsPt: [100, 100] }).error).toMatch(/merged cells/);
    });

    test('toTablePatch carries styleOps and opCount counts them', () => {
        const model = createTableModel(REGION);
        model.setCell(2, 1, 'new a');
        model.setTableStyle({ style: 'TableGrid' });
        model.setBorders({ borders: { all: 'single' } });
        expect(model.opCount).toBe(3);
        const patch = model.toTablePatch();
        expect(patch.cells).toHaveLength(1);
        expect(patch.styleOps).toHaveLength(2);
    });

    test('executeTableTool dispatches the style tools', () => {
        const model = createTableModel(REGION);
        expect(executeTableTool(model, 'set_table_style', { style: 'TableGrid' }).ok).toBe(true);
        expect(executeTableTool(model, 'set_borders', { borders: { all: 'none' } }).ok).toBe(true);
        expect(executeTableTool(model, 'set_cell_format', { row: 1, shadingColor: 'red' }).ok).toBe(true);
        expect(executeTableTool(model, 'set_font', { row: 1, font: { bold: true } }).ok).toBe(true);
        expect(executeTableTool(model, 'set_header_row', {}).ok).toBe(true);
        expect(executeTableTool(model, 'set_layout', { alignment: 'centered' }).ok).toBe(true);
        expect(executeTableTool(model, 'set_column_widths', { widthsPt: [100, 100] }).ok).toBe(true);
        expect(executeTableTool(model, 'set_style', {}).ok).toBe(false);
    });
});

describe('multi-table sessions', () => {
    const REGION_TWO = {
        rowCount: 2,
        colCount: 3,
        values: [
            ['x1', 'x2', 'x3'],
            ['y1', 'y2', 'y3'],
        ],
        bounds: { startRow: 1, endRow: 2, startCol: 1, endCol: 3 },
        merged: false,
        shadowKeys: new Set(),
    };

    test('get_state lists every table grid under its table index', () => {
        const model = createTableModel([REGION, REGION_TWO]);
        const { result } = model.getState();
        expect(result.tableCount).toBe(2);
        expect(result.grid).toContain('table 1 (3x2):');
        expect(result.grid).toContain('table 2 (2x3):');
        expect(result.grid).toContain('[R1C1] Header A');
        expect(result.grid).toContain('[R2C3] y3');
    });

    test('tools address tables by tableIndex and default to table 1', () => {
        const model = createTableModel([REGION, REGION_TWO]);
        expect(model.setCell(2, 1, 'new a').ok).toBe(true); // table 1 default
        expect(model.setCell(1, 1, 'X1', 2).ok).toBe(true);
        expect(model.deleteRow(2, 2).ok).toBe(true);
        const patch = model.toTablePatch();
        expect(patch.cells).toEqual([
            { tableIndex: 1, row: 2, col: 1, text: 'new a' },
            { tableIndex: 2, row: 1, col: 1, text: 'X1' },
        ]);
        expect(patch.rowOps).toEqual([{ tableIndex: 2, op: 'delete', row: 2 }]);
        expect(patch.tableCount).toBe(2);
    });

    test('unknown tableIndex is rejected constructively', () => {
        const model = createTableModel([REGION, REGION_TWO]);
        expect(model.setCell(1, 1, 'z', 3).error).toMatch(/"tableIndex" 3/);
        expect(model.setTableStyle({ tableIndex: 9, style: 'TableGrid' }).error).toMatch(/"tableIndex" 9/);
    });

    test('style ops land on the right table; merged-only table blocks its own row ops', () => {
        const model = createTableModel([
            REGION,
            { ...REGION_TWO, merged: true, shadowKeys: new Set(['1,2']) },
        ]);
        expect(model.setTableStyle({ tableIndex: 2, style: 'GridTable2', bandedRows: true }).ok).toBe(true);
        expect(model.setBorders({ tableIndex: 1, borders: { all: 'none' } }).ok).toBe(true);
        expect(model.deleteRow(2, 2).error).toMatch(/merged cells/);
        expect(model.deleteRow(2, 1).ok).toBe(true);
        const patch = model.toTablePatch();
        expect(patch.styleOps).toEqual([
            { type: 'borders', tool: 'set_borders', tableIndex: 1, borders: { top: { type: 'none' }, bottom: { type: 'none' }, left: { type: 'none' }, right: { type: 'none' }, insideH: { type: 'none' }, insideV: { type: 'none' } } },
            { type: 'tableStyle', tool: 'set_table_style', tableIndex: 2, style: 'GridTable2', bandedRows: true },
        ]);
    });

    test('get_state pendingOps carry tableIndex per op', () => {
        const model = createTableModel([REGION, REGION_TWO]);
        model.setCell(1, 1, 'tip', 2);
        model.deleteRow(3, 1);
        const { result } = model.getState();
        expect(result.pendingOps).toContainEqual({ tool: 'set_cell', tableIndex: 2, row: 1, col: 1, text: 'tip' });
        expect(result.pendingOps).toContainEqual({ op: 'delete', row: 3, tableIndex: 1 });
    });

    test('opCount aggregates across tables', () => {
        const model = createTableModel([REGION, REGION_TWO]);
        model.setCell(1, 1, 'a', 2);
        model.setTableStyle({ tableIndex: 1, style: 'TableGrid' });
        expect(model.opCount).toBe(2);
    });

    test('executes the same tool dispatch against any table', () => {
        const model = createTableModel([REGION, REGION_TWO]);
        expect(executeTableTool(model, 'set_cell', { tableIndex: 2, row: 2, col: 2, text: 'Y2' }).ok).toBe(true);
        expect(executeTableTool(model, 'set_column_widths', { tableIndex: 2, widthsPt: [100, 100, 100] }).ok).toBe(true);
        expect(model.toTablePatch().styleOps[0].tableIndex).toBe(2);
    });
});
