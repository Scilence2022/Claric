/**
 * Specs for src/lib/selection-context.js : formatTableMarkdown +
 * formatMixedContext.
 *
 * Covers the pure formatting of table selections into LLM-ready markdown:
 * grid rendering (marker, header, separator, data rows), cell sanitization
 * (pipes, in-cell newlines), ragged-row padding, token-budget truncation
 * (data rows first, then columns, hard clip last), truncation notes, and
 * the document-ordered interleaving of mixed paragraph+table selections.
 */

const {
    formatTableMarkdown,
    formatMixedContext,
    formatCursorContext,
} = require('../src/lib/selection-context.js');

describe('formatTableMarkdown', () => {
    it('renders marker, header, separator, and data rows', () => {
        const out = formatTableMarkdown([
            ['Name', 'Qty'],
            ['Apple', '3'],
            ['Pear', '5'],
        ]);
        expect(out).toBe(
            '[TABLE 3 rows × 2 cols]\n' +
            '| Name | Qty |\n' +
            '| --- | --- |\n' +
            '| Apple | 3 |\n' +
            '| Pear | 5 |'
        );
    });

    it('returns empty string for empty or non-array input', () => {
        expect(formatTableMarkdown([])).toBe('');
        expect(formatTableMarkdown(null)).toBe('');
        expect(formatTableMarkdown([[]])).toBe('');
        expect(formatTableMarkdown('nope')).toBe('');
    });

    it('escapes pipes so cell content cannot forge columns', () => {
        const out = formatTableMarkdown([['a|b', 'c'], ['d', 'e']]);
        expect(out.split('\n')[1]).toBe('| a\\|b | c |');
    });

    it('converts in-cell newlines to <br> and trims cell text', () => {
        const out = formatTableMarkdown([['line1\nline2', '  x  ']]);
        expect(out.split('\n')[1]).toBe('| line1<br>line2 | x |');
    });

    it('normalizes CR/LF variants inside cells', () => {
        const out = formatTableMarkdown([['a\r\nb']]);
        expect(out).toContain('a<br>b');
    });

    it('pads ragged rows to the widest row', () => {
        const out = formatTableMarkdown([['a', 'b', 'c'], ['d']]);
        expect(out.split('\n')[3]).toBe('| d |  |  |');
    });

    it('keeps empty cells as empty cells', () => {
        const out = formatTableMarkdown([['a', ''], ['', 'b']]);
        expect(out.split('\n')[3]).toBe('|  | b |');
    });

    it('keeps a header-only table intact without truncation notes', () => {
        const out = formatTableMarkdown([['a', 'b']]);
        expect(out).toBe(
            '[TABLE 1 rows × 2 cols]\n' +
            '| a | b |\n' +
            '| --- | --- |'
        );
    });

    it('appends the selection note to the marker line', () => {
        const out = formatTableMarkdown([['a'], ['b']], { note: 'user selected R2C1–R2C1' });
        expect(out.startsWith('[TABLE 2 rows × 1 cols — user selected R2C1–R2C1]')).toBe(true);
    });

    it('truncates data rows from the end under the budget and notes it', () => {
        const values = [
            ['H1', 'H2'],
            ['r1c1', 'r1c2'],
            ['r2c1', 'r2c2'],
            ['r3c1', 'r3c2'],
        ];
        // Full render is 97 chars; maxTokens 21 → 84-char budget, so the
        // third data row drops. Header and separator always survive.
        const out = formatTableMarkdown(values, { maxTokens: 21 });
        expect(out).toContain('| H1 | H2 |');
        expect(out).toContain('| r2c1 | r2c2 |');
        expect(out).not.toContain('| r3c1 | r3c2 |');
        expect(out).toContain('[... 1 more row(s) truncated]');
        expect(out).not.toContain('column(s) truncated');
    });

    it('drops columns from the right when the header alone busts the budget', () => {
        const a = 'a'.repeat(30);
        const b = 'b'.repeat(30);
        // 2-column header render is 105 chars; 1-column is 65. maxTokens 17
        // → 68-char budget: the second column drops, and no data row fits.
        const out = formatTableMarkdown([[a, b], ['c', 'd']], { maxTokens: 17 });
        expect(out).toContain(`| ${a} |`);
        expect(out).not.toContain(b);
        expect(out).toContain('[... 1 more row(s) truncated]');
        expect(out).toContain('[... 1 more column(s) truncated]');
    });

    it('hard-clips a pathological single giant cell and notes it', () => {
        const out = formatTableMarkdown([['x'.repeat(100)]], { maxTokens: 5 });
        expect(out.split('\n[... truncated]')[0].length).toBeLessThanOrEqual(20);
        expect(out).toContain('[... truncated]');
    });
});

describe('formatMixedContext', () => {
    it('interleaves paragraphs and table blocks in document order', () => {
        const out = formatMixedContext([
            { type: 'paragraph', text: 'Caption above.' },
            { type: 'table', values: [['h1', 'h2'], ['a', 'b']] },
            { type: 'paragraph', text: 'Note below.' },
        ]);
        expect(out).toBe(
            'Caption above.\n\n' +
            '[TABLE 2 rows × 2 cols]\n' +
            '| h1 | h2 |\n' +
            '| --- | --- |\n' +
            '| a | b |\n\n' +
            'Note below.'
        );
    });

    it('joins consecutive paragraphs one per line and drops blanks', () => {
        const out = formatMixedContext([
            { type: 'paragraph', text: 'one' },
            { type: 'paragraph', text: '' },
            { type: 'paragraph', text: 'two' },
        ]);
        expect(out).toBe('one\ntwo');
    });

    it('renders multiple tables as separate blocks', () => {
        const out = formatMixedContext([
            { type: 'table', values: [['a']] },
            { type: 'paragraph', text: 'between' },
            { type: 'table', values: [['b']] },
        ]);
        expect(out).toBe(
            '[TABLE 1 rows × 1 cols]\n| a |\n| --- |\n\n' +
            'between\n\n' +
            '[TABLE 1 rows × 1 cols]\n| b |\n| --- |'
        );
    });

    it('passes the per-table note through to the marker', () => {
        const out = formatMixedContext([
            { type: 'table', values: [['a']], note: 'merged cells' },
        ]);
        expect(out).toContain('[TABLE 1 rows × 1 cols — merged cells]');
    });

    it('returns empty string for empty parts', () => {
        expect(formatMixedContext([])).toBe('');
        expect(formatMixedContext(null)).toBe('');
    });
});

describe('formatCursorContext', () => {
    it('renders heading and cursor paragraph', () => {
        const out = formatCursorContext({
            paragraphText: 'Payment shall be made within 30 days.',
            headingText: '3.2 Payment Terms',
            headingLevel: 2,
        });
        expect(out).toBe(
            'Nearest section: "3.2 Payment Terms" (heading level 2)\n' +
            'Cursor paragraph: "Payment shall be made within 30 days."'
        );
    });

    it('renders heading-only when the caret sits in an empty paragraph', () => {
        const out = formatCursorContext({ paragraphText: '', headingText: 'Intro', headingLevel: 1 });
        expect(out).toBe('Nearest section: "Intro" (heading level 1)');
    });

    it('renders paragraph-only when no heading was found', () => {
        const out = formatCursorContext({ paragraphText: 'standalone text' });
        expect(out).toBe('Cursor paragraph: "standalone text"');
    });

    it('returns empty string when neither paragraph nor heading is available', () => {
        expect(formatCursorContext({})).toBe('');
        expect(formatCursorContext()).toBe('');
        expect(formatCursorContext({ paragraphText: '   ', headingText: '' })).toBe('');
    });

    it('clips a long cursor paragraph at 300 chars and heading at 120 chars', () => {
        const para = 'x'.repeat(400);
        const heading = 'h'.repeat(200);
        const out = formatCursorContext({ paragraphText: para, headingText: heading, headingLevel: 1 });
        expect(out).toContain(`"${'h'.repeat(120)}"`);
        expect(out).toContain(`"${'x'.repeat(300)}"`);
        expect(out).not.toContain('x'.repeat(301));
    });

    it('notes when the caret is inside a table cell', () => {
        const out = formatCursorContext({ paragraphText: 'cell text', inTable: true });
        expect(out).toContain('The cursor is inside a table cell.');
    });

    it('normalizes CR in paragraph text', () => {
        const out = formatCursorContext({ paragraphText: 'a\r\nb' });
        expect(out).toContain('a\nb');
    });
});
