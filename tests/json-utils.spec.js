/**
 * Shared JSON extraction utilities (json-utils) tests.
 *
 * The critical regression covered here: trailing-comma cleanup used to be a
 * blind `,\s*([}\]])` regex applied to the whole JSON slice, which also
 * fired INSIDE string literals — a cell/tool-arg value like
 * "items: a, b, ]" silently lost its comma before the text reached the
 * document. Cleanup must be string-aware and must never run on valid JSON.
 */

const {
    stripTrailingCommas,
    balancedJsonCandidates,
    extractJsonObject,
    extractJsonArray,
} = require('../src/lib/json-utils.js');

describe('stripTrailingCommas', () => {
    test('removes trailing commas outside strings', () => {
        expect(stripTrailingCommas('{"a":1,}')).toBe('{"a":1}');
        expect(stripTrailingCommas('[1,2,3,]')).toBe('[1,2,3]');
        expect(stripTrailingCommas('{\n  "a": 1,\n  "b": 2,\n}')).toBe('{\n  "a": 1,\n  "b": 2\n}');
    });

    test('never touches commas inside string values (the cell-text regression)', () => {
        const json = '{"cells":[{"row":1,"col":2,"text":"items: a, b, ] and {, }"}]}';
        expect(stripTrailingCommas(json)).toBe(json);
        expect(JSON.parse(stripTrailingCommas(json))).toEqual({
            cells: [{ row: 1, col: 2, text: 'items: a, b, ] and {, }' }],
        });
    });

    test('handles escaped quotes and mixed content', () => {
        const json = '{"text":"she said \\"hi\\", then left",}';
        expect(JSON.parse(stripTrailingCommas(json)).text).toBe('she said "hi", then left');
        // A real trailing comma after a string that itself ends with a comma.
        const tricky = '{"a":"x,",}';
        expect(JSON.parse(stripTrailingCommas(tricky)).a).toBe('x,');
    });
});

describe('balancedJsonCandidates', () => {
    test('extracts complete containers from surrounding prose', () => {
        const cands = balancedJsonCandidates('Sure! {"tool":"set_cell"} — hope that helps.');
        expect(cands).toHaveLength(1);
        expect(JSON.parse(cands[0].text)).toEqual({ tool: 'set_cell' });
    });

    test('ignores braces inside strings and extracts multiple candidates', () => {
        const cands = balancedJsonCandidates('{"a":"{ not a container"} trailing [1,2]');
        expect(cands.map((c) => JSON.parse(c.text))).toEqual([{ a: '{ not a container' }, [1, 2]]);
    });

    test('terminates a candidate at a mismatched closing', () => {
        const cands = balancedJsonCandidates('{"a":[1}}');
        expect(cands).toHaveLength(1);
        expect(cands[0].text).toBe('{"a":[1}');
    });
});

describe('extractJsonObject', () => {
    test('parses a plain, fenced, and prose-wrapped object', () => {
        expect(extractJsonObject('{"tool":"finish"}')).toEqual({ tool: 'finish' });
        expect(extractJsonObject('```json\n{"tool":"finish"}\n```')).toEqual({ tool: 'finish' });
        expect(extractJsonObject('Here you go: {"tool":"finish"} thanks')).toEqual({ tool: 'finish' });
    });

    test('recovers trailing commas and string content with ", ]" survives', () => {
        expect(extractJsonObject('{"cells":[{"text":"a, b, }"}],}'))
            .toEqual({ cells: [{ text: 'a, b, }' }] });
    });

    test('skips prose brace pairs and still finds the real object', () => {
        const raw = 'Note {like this} the patch: {"tool":"set_cell","args":{"text":"x"}}';
        expect(extractJsonObject(raw)).toEqual({ tool: 'set_cell', args: { text: 'x' } });
    });

    test('array reply is a "not an object" failure, primitives and empty are no-object', () => {
        expect(() => extractJsonObject('[{"a":1}]')).toThrow(/not an object/);
        expect(() => extractJsonObject('123')).toThrow(/no JSON object/);
        expect(() => extractJsonObject('   ')).toThrow(/no JSON object/);
        expect(() => extractJsonObject('no json here')).toThrow(/no JSON object/);
    });

    test('custom messages keep each protocol layer error contract', () => {
        expect(() => extractJsonObject('nope', {
            noObjectMessage: 'Table patch response contains no JSON object',
            parseFailedPrefix: 'Table patch JSON parse failed: ',
        })).toThrow('Table patch response contains no JSON object');
        expect(() => extractJsonObject('{"a":', {
            noObjectMessage: 'Table patch response contains no JSON object',
            parseFailedPrefix: 'Table patch JSON parse failed: ',
        })).toThrow(/Table patch JSON parse failed: /);
    });
});

describe('extractJsonArray', () => {
    test('parses plain, fenced, and prose-wrapped arrays', () => {
        expect(extractJsonArray('[{"type":"edit"}]')).toEqual({ value: [{ type: 'edit' }], error: null });
        expect(extractJsonArray('```json\n[1,2,]\n```')).toEqual({ value: [1, 2], error: null });
        expect(extractJsonArray('Plan: [{"type":"qa"}] done')).toEqual({ value: [{ type: 'qa' }], error: null });
    });

    test('object/primitive replies and empty input report no array', () => {
        expect(extractJsonArray('{"type":"edit"}').error).toMatch(/no JSON array/);
        expect(extractJsonArray('42').error).toMatch(/no JSON array/);
        expect(extractJsonArray('').error).toMatch(/no JSON array/);
    });

    test('malformed container reports not-valid-JSON with the parse reason', () => {
        const { value, error } = extractJsonArray('[{"type": oops}]');
        expect(value).toBeNull();
        expect(error).toMatch(/not valid JSON/);
    });
});
