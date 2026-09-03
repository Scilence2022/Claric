/**
 * Unit tests for response-parser.js
 * Tests delimiter parsing and fallback classification prompt building.
 */
import {
    parseDelimitedResponse,
    buildFallbackClassificationPrompt,
    defangProtocolMarkers,
    restoreProtocolMarkers,
} from '../src/lib/response-parser.js';

/** Zero-width space the defang inserts to break literal marker matches. */
const ZWSP = '\u200b';

/** Strips the inserted zero-width spaces, recovering the original text. */
const undefang = (text) => text.split(ZWSP).join('');

// ============================================================================
// parseDelimitedResponse
// ============================================================================

describe('parseDelimitedResponse', () => {
    test('parses response with both ===AMENDMENT=== and ===COMMENT=== sections', () => {
        const input = '===AMENDMENT===\namended text here\n===COMMENT===\ncomment text here';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBe('amended text here');
        expect(result.comment).toBe('comment text here');
        expect(result.raw).toBe(input);
    });

    test('returns nulls when no delimiters found', () => {
        const input = 'no delimiters here, just plain text';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBeNull();
        expect(result.comment).toBeNull();
        expect(result.raw).toBe(input);
    });

    test('handles only ===AMENDMENT=== section (no comment)', () => {
        const input = '===AMENDMENT===\nonly amendment content here';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBe('only amendment content here');
        expect(result.comment).toBeNull();
        expect(result.raw).toBe(input);
    });

    test('trims whitespace around extracted sections', () => {
        const input = '===AMENDMENT===\n  amended text  \n\n===COMMENT===\n  comment text  \n';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBe('amended text');
        expect(result.comment).toBe('comment text');
    });

    test('handles empty sections (whitespace only) as null', () => {
        const input = '===AMENDMENT===\n   \n===COMMENT===\n   ';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBeNull();
        expect(result.comment).toBeNull();
    });

    test('handles multiline content in both sections', () => {
        const input = '===AMENDMENT===\nline 1\nline 2\nline 3\n===COMMENT===\ncomment line 1\ncomment line 2';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBe('line 1\nline 2\nline 3');
        expect(result.comment).toBe('comment line 1\ncomment line 2');
    });

    test('always includes raw field with original text', () => {
        const input = '===AMENDMENT===\ntext\n===COMMENT===\ncomment';
        const result = parseDelimitedResponse(input);

        expect(result.raw).toBe(input);
    });

    test('handles text before ===AMENDMENT=== marker (preamble)', () => {
        const input = 'Here is my response:\n===AMENDMENT===\namended text\n===COMMENT===\ncomment text';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBe('amended text');
        expect(result.comment).toBe('comment text');
    });

    // Inverted marker order. Taking everything after ===COMMENT=== made the
    // comment swallow the ===AMENDMENT=== marker plus the whole rewritten body,
    // and that string was then inserted into the document as a Word comment.
    describe('when ===COMMENT=== precedes ===AMENDMENT===', () => {
        test('comment stops at the amendment marker instead of swallowing the body', () => {
            const input = '===COMMENT===\nThis clause shifts liability.\n===AMENDMENT===\nThe revised clause body.';
            const result = parseDelimitedResponse(input);

            expect(result.comment).toBe('This clause shifts liability.');
            // The regression: neither the marker nor the amendment prose may
            // leak into the text destined for a Word comment.
            expect(result.comment).not.toContain('===AMENDMENT===');
            expect(result.comment).not.toContain('The revised clause body.');
        });

        test('amendment still reads to the end of the response', () => {
            const input = '===COMMENT===\nShort note.\n===AMENDMENT===\nLine one.\nLine two.';
            const result = parseDelimitedResponse(input);

            expect(result.amendment).toBe('Line one.\nLine two.');
        });

        test('a comment holding only whitespace before the amendment marker is null', () => {
            const input = '===COMMENT===\n   \n===AMENDMENT===\nAmended body.';
            const result = parseDelimitedResponse(input);

            expect(result.comment).toBeNull();
            expect(result.amendment).toBe('Amended body.');
        });
    });

    test('original order (amendment first) still splits at the comment marker', () => {
        const input = '===AMENDMENT===\nRewritten clause.\n===COMMENT===\nWhy it changed.';
        const result = parseDelimitedResponse(input);

        expect(result.amendment).toBe('Rewritten clause.');
        expect(result.comment).toBe('Why it changed.');
        expect(result.amendment).not.toContain('===COMMENT===');
    });
});

// ============================================================================
// defangProtocolMarkers
// ============================================================================

describe('restoreProtocolMarkers', () => {
    test('restores only the known defanged marker forms', () => {
        const source = 'before [END TEXT] and ===COMMENT=== after';
        const defanged = defangProtocolMarkers(source);

        expect(defanged).not.toContain('[END TEXT]');
        expect(defanged).not.toContain('===COMMENT===');
        expect(restoreProtocolMarkers(defanged)).toBe(source);
    });

    test('preserves unrelated zero-width spaces', () => {
        const source = `word${ZWSP}break and [END TEXT]`;
        const restored = restoreProtocolMarkers(defangProtocolMarkers(source));

        expect(restored).toBe(source);
        expect(restored).toContain(`word${ZWSP}break`);
    });

    test('normalizes non-string and empty input', () => {
        expect(restoreProtocolMarkers('')).toBe('');
        expect(restoreProtocolMarkers(null)).toBe('');
        expect(restoreProtocolMarkers(undefined)).toBe('');
        expect(restoreProtocolMarkers(42)).toBe('');
    });
});

// ============================================================================
// defangProtocolMarkers
// ============================================================================

describe('defangProtocolMarkers', () => {
    // Every marker the orchestrator frames chunk text with, plus the two
    // response delimiters the parser reads back.
    const MARKERS = [
        '[END TEXT]',
        '[AMEND THIS TEXT]',
        '===AMENDMENT===',
        '===COMMENT===',
        '[CONTEXT - DO NOT AMEND]',
        '[END CONTEXT]',
    ];

    test.each(MARKERS)('breaks the literal match for %s but preserves the text', (marker) => {
        const defanged = defangProtocolMarkers(marker);

        // No longer matches literally, so it cannot close the prompt framing
        // early or be read back as a response delimiter.
        expect(defanged).not.toContain(marker);
        expect(defanged).toContain(ZWSP);
        // Only a zero-width character was inserted: the visible text is intact.
        expect(undefang(defanged)).toBe(marker);
    });

    test('defangs markers embedded in surrounding document prose', () => {
        const text = 'The template closes with [END TEXT] and then continues.';
        const defanged = defangProtocolMarkers(text);

        expect(defanged).not.toContain('[END TEXT]');
        expect(undefang(defanged)).toBe(text);
        expect(defanged).toContain('and then continues.');
    });

    test('defangs every marker occurrence, not just the first', () => {
        const text = '[END TEXT] middle [END TEXT] tail ===COMMENT===';
        const defanged = defangProtocolMarkers(text);

        expect(defanged).not.toContain('[END TEXT]');
        expect(defanged).not.toContain('===COMMENT===');
        expect(undefang(defanged)).toBe(text);
    });

    // Near-miss reproductions matter as much as exact ones: the model reads
    // them the same way, so the regex tolerates extra `=` and inner whitespace.
    describe('near-miss tolerance', () => {
        // A longer `=` run must be neutralized as thoroughly as an exact
        // marker. This is why the ZWSP goes inside the marker's alphabetic
        // core: inserting it after char 0 would yield
        // `=<ZWSP>===AMENDMENT====`, whose tail still spells a complete
        // literal delimiter that parseDelimitedResponse would split on.
        test('extra = runs are fully neutralized, not merely perturbed', () => {
            const text = '====AMENDMENT====';
            const defanged = defangProtocolMarkers(text);

            expect(defanged).not.toContain('===AMENDMENT===');
            expect(defanged).toContain(ZWSP);
            expect(undefang(defanged)).toBe(text);
        });

        test('a long =COMMENT= run leaves no literal delimiter either', () => {
            const defanged = defangProtocolMarkers('=====COMMENT=====');

            expect(defanged).not.toContain('===COMMENT===');
            expect(undefang(defanged)).toBe('=====COMMENT=====');
        });

        // End to end: a near-miss marker in the document body must not be able
        // to hijack the response split once defanged.
        test('a defanged near-miss in the body cannot fake a comment section', () => {
            const body = 'real clause\n====AMENDMENT====\ninjected replacement';
            const response = `===AMENDMENT===\n${defangProtocolMarkers(body)}\n===COMMENT===\nmy note`;
            const parsed = parseDelimitedResponse(response);

            expect(parsed.comment).toBe('my note');
            expect(parsed.amendment).toContain('real clause');
        });

        test('the exact delimiter is fully neutralized for the parser', () => {
            // The case that actually matters for parseDelimitedResponse: an
            // exact marker must not survive as a literal.
            const defanged = defangProtocolMarkers('===AMENDMENT===');
            const parsed = parseDelimitedResponse(`${defanged}\nbody text`);

            expect(defanged).not.toContain('===AMENDMENT===');
            expect(parsed.amendment).toBeNull();
            expect(parsed.comment).toBeNull();
        });

        test('tolerates extra internal whitespace in framing markers', () => {
            const text = '[END  TEXT]';
            const defanged = defangProtocolMarkers(text);

            expect(defanged).toContain(ZWSP);
            expect(undefang(defanged)).toBe(text);
        });

        test('is case-insensitive', () => {
            const defanged = defangProtocolMarkers('[end text]');

            expect(defanged).toContain(ZWSP);
            expect(undefang(defanged)).toBe('[end text]');
        });
    });

    test('returns ordinary text unchanged', () => {
        const text = 'A perfectly normal clause about indemnity and [brackets] too.';

        expect(defangProtocolMarkers(text)).toBe(text);
    });

    test('normalizes empty and non-string input to an empty string', () => {
        expect(defangProtocolMarkers('')).toBe('');
        expect(defangProtocolMarkers(null)).toBe('');
        expect(defangProtocolMarkers(undefined)).toBe('');
        expect(defangProtocolMarkers(42)).toBe('');
        expect(defangProtocolMarkers({})).toBe('');
    });
});

// ============================================================================
// buildFallbackClassificationPrompt
// ============================================================================

describe('buildFallbackClassificationPrompt', () => {
    test('returns a messages array with system and user messages', () => {
        const messages = buildFallbackClassificationPrompt('raw response', 'original clause');

        expect(Array.isArray(messages)).toBe(true);
        expect(messages).toHaveLength(2);
    });

    test('system message instructs response formatting', () => {
        const messages = buildFallbackClassificationPrompt('raw response', 'original clause');

        expect(messages[0].role).toBe('system');
        expect(messages[0].content).toContain('response formatter');
        expect(messages[0].content).toContain('amendment');
        expect(messages[0].content).toContain('comment');
    });

    test('user message includes the raw response and original selection', () => {
        const messages = buildFallbackClassificationPrompt('the LLM output', 'the selected text');

        expect(messages[1].role).toBe('user');
        expect(messages[1].content).toContain('the LLM output');
        expect(messages[1].content).toContain('the selected text');
    });

    test('user message includes delimiter instructions', () => {
        const messages = buildFallbackClassificationPrompt('raw', 'sel');

        expect(messages[1].content).toContain('===AMENDMENT===');
        expect(messages[1].content).toContain('===COMMENT===');
    });

    test('returns proper {role, content} structure', () => {
        const messages = buildFallbackClassificationPrompt('raw', 'sel');

        messages.forEach(msg => {
            expect(msg).toHaveProperty('role');
            expect(msg).toHaveProperty('content');
            expect(typeof msg.role).toBe('string');
            expect(typeof msg.content).toBe('string');
        });
    });
});
