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
            'get_state', 'set_cell', 'insert_row', 'delete_row',
        ]);
    });
});
