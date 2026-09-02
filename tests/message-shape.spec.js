/**
 * Message shape tests (src/taskpane/message-shape.js).
 *
 * The persisted chat message had two independent normalizers — one in
 * sessions.js (save leg) and one in ui/chat-view.js (load leg) — validating
 * the same 11 fields. A field added to one and forgotten in the other was
 * silently dropped on that leg of the round-trip. Both now normalize through
 * this module; these tests pin the shared shape and the round-trip identity.
 */

const { newId, normalizeMessage, normalizeAttachments, normalizeCitations } =
    require('../src/taskpane/message-shape.js');
const { __testing } = require('../src/taskpane/sessions.js');

describe('newId', () => {
    test('prefixes the id and stays unique', () => {
        const a = newId('m');
        const b = newId('m');
        expect(a.startsWith('m-')).toBe(true);
        expect(a).not.toBe(b);
    });
});

describe('normalizeMessage', () => {
    test('fills every field for an empty input', () => {
        const m = normalizeMessage(undefined);
        expect(m.role).toBe('user');
        expect(m.text).toBe('');
        expect(m.status).toBe('');
        expect(m.error).toBeNull();
        expect(m.worklog).toBeNull();
        expect(m.model).toBeNull();
        expect(m.citations).toEqual([]);
        expect(m.proposals).toEqual([]);
        expect(m.attachments).toEqual([]);
        expect(typeof m.id).toBe('string');
        expect(typeof m.ts).toBe('string');
    });

    test('drops unknown fields and coerces wrong types', () => {
        const m = normalizeMessage({
            id: 'm-1', role: 'assistant', text: 'hi', status: 7, error: 42,
            worklog: { count: '3', durationMs: 'x' },
            model: { sections: '2' },
            secret: 'should not survive',
            ts: '2026-01-01T00:00:00.000Z',
        });
        expect(m.status).toBe('');
        expect(m.error).toBeNull();
        expect(m.worklog).toEqual({ count: 3, durationMs: 0 });
        expect(m.model).toEqual({ sections: 2 });
        expect(m.secret).toBeUndefined();
    });

    test('round-trips a normalized message unchanged', () => {
        const once = normalizeMessage({
            id: 'm-1', role: 'assistant', text: 'answer', status: 'Done', error: null,
            worklog: { count: 2, durationMs: 1500 },
            model: { sections: 1 },
            citations: [{ label: 'Clause 1', searchText: 'Clause 1 term' }],
            proposals: [{ title: 'T', state: 'applied' }],
            attachments: [{ name: 'a.txt', kind: 'text', size: 5 }],
            ts: '2026-01-01T00:00:00.000Z',
        });
        expect(normalizeMessage(once)).toEqual(once);
    });

    test("sessions.js's save leg is the same normalizer", () => {
        const input = {
            id: 'm-1', role: 'assistant', text: 'a',
            citations: [{ label: 'L', searchText: 'S' }],
            attachments: [{ name: 'p.png', kind: 'image', size: 9, dataUrl: 'data:image/png;base64,zz' }],
            ts: '2026-01-01T00:00:00.000Z',
        };
        expect(__testing.stripMessage(input)).toEqual(normalizeMessage(input));
    });
});

describe('normalizeAttachments', () => {
    test('keeps display metadata only', () => {
        expect(normalizeAttachments([
            { name: 'a.txt', kind: 'text', size: 5, text: 'body' },
            { name: 'p.png', kind: 'image', size: 9, dataUrl: 'data:image/png;base64,zz' },
        ])).toEqual([
            { name: 'a.txt', kind: 'text', size: 5 },
            { name: 'p.png', kind: 'image', size: 9 },
        ]);
    });

    test('drops nameless and non-object entries', () => {
        expect(normalizeAttachments([null, 'x', { size: 5 }, { name: '' }])).toEqual([]);
        expect(normalizeAttachments(null)).toEqual([]);
    });
});

describe('normalizeCitations', () => {
    test('keeps label/searchText as strings', () => {
        expect(normalizeCitations([{ label: 'Clause 5', searchText: 'Clause 5 payment' }]))
            .toEqual([{ label: 'Clause 5', searchText: 'Clause 5 payment' }]);
        expect(normalizeCitations([{ label: 7 }])).toEqual([{ label: '7', searchText: '' }]);
    });

    test('tolerates junk', () => {
        expect(normalizeCitations(null)).toEqual([]);
        expect(normalizeCitations([null, 'x'])).toEqual([]);
    });
});
