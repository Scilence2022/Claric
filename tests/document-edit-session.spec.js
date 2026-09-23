import { runDocumentEditSession } from '../src/lib/document-edit-session.js';

const snapshot = { id: 's', blocks: [
    { id: 'p1', text: 'Discussion', headingLevel: 1 },
    { id: 'p2', text: 'Mechanism.' },
    { id: 'p3', text: 'Limitations.' },
] };
const call = (tool, args = {}) => JSON.stringify({ tool, args });
const contract = call('set_edit_contract', { goal: 'Integrate XXX', requirements: ['Discuss XXX before limitations'] });
const read = call('read_blocks', { ids: ['p2', 'p3'] });
const insert = call('stage_patch', { operations: [{ kind: 'insert', afterId: 'p2', paragraphs: ['XXX discussion.'], reason: 'Between mechanism and limitations.' }] });
const review = (overrides = {}) => JSON.stringify({ satisfied: true, instructionSatisfied: true, summary: 'Suitable placement.', checks: [{ requirement: 0, satisfied: true, evidence: 'Draft paragraph before p3.' }], issues: [], ...overrides });
const finish = call('finish', { summary: 'Proposed XXX discussion before limitations.' });

function scripted(replies) {
    const send = jest.fn(async () => {
        if (!replies.length) throw new Error('Unexpected request');
        const item = replies.shift();
        if (item instanceof Error) throw item;
        return item;
    });
    return send;
}

test('completes insertion, draft inspection and independent review before returning a proposal', async () => {
    const send = scripted([call('read_outline'), call('search_document', { query: 'Mechanism' }), contract, read, insert,
        call('read_draft'), call('validate_patch'), review(), finish]);
    const result = await runDocumentEditSession({ snapshot, instruction: 'Insert XXX at an appropriate place, keeping headings.', send });
    expect(result.status).toBe('staged');
    expect(result.patch.changes[0]).toMatchObject({ afterId: 'p2', beforeId: 'p3' });
    const reviewRequest = send.mock.calls.find(([messages]) => messages[0].content.startsWith('Review a PROPOSED'))[0];
    expect(reviewRequest[1].content).toContain('keeping headings');
    expect(reviewRequest[1].content).toContain('Limitations.');
    expect(JSON.parse(reviewRequest[1].content).documentEvidence).toMatchObject([
        { id: 'p2', text: 'Mechanism.' }, { id: 'p3', text: 'Limitations.' },
    ]);
});

test('a prose insertion can be formatted before the same draft is reviewed', async () => {
    const send = scripted([contract, read, insert, call('read_draft'),
        call('stage_patch', { operations: [{ kind: 'format_new', blockId: 'draft-1', format: { bold: true }, reason: 'Requested emphasis.' }] }),
        call('read_draft'), call('validate_patch'), review(), finish]);
    const result = await runDocumentEditSession({ snapshot, instruction: 'Insert XXX and bold the new paragraph.', send });
    expect(result.patch.changes[0].paragraphFormats).toEqual([{ bold: true }]);
    const reviewRequest = send.mock.calls.find(([messages]) => messages[0].content.startsWith('Review a PROPOSED'))[0];
    expect(reviewRequest[1].content).toContain('"bold":true');
});

test('premature finish and invalid arguments get observations; failed semantic review can be repaired', async () => {
    const send = scripted([finish, contract, call('stage_patch', { operations: [] }), read, insert,
        call('validate_patch'), call('read_draft'), call('validate_patch'), review({ satisfied: false, issues: ['Transition is abrupt.'] }),
        call('stage_patch', { operations: [{ kind: 'replace', blockId: 'p3', text: 'Nevertheless, limitations remain.', reason: 'Transition.' }] }),
        finish, call('read_draft'), call('validate_patch'), review(), finish]);
    const result = await runDocumentEditSession({ snapshot, instruction: 'Insert XXX and connect the discussion.', send });
    expect(result.patch.changes).toHaveLength(2);
    expect(result.review.ok).toBe(true);
    expect(send.mock.calls.some(([messages]) => JSON.stringify(messages).includes('Read the latest draft'))).toBe(true);
    expect(send.mock.calls.some(([messages]) => JSON.stringify(messages).includes('Transition is abrupt'))).toBe(true);
});

test('no-op requires explicit evidence that the original already satisfies the request', async () => {
    const send = scripted([contract, read, call('read_draft'), call('validate_patch'), review({ noChangeNeeded: true }), finish]);
    const result = await runDocumentEditSession({ snapshot, instruction: 'Ensure limitations are mentioned.', send });
    expect(result.status).toBe('no_op');
    expect(result.patch.changes).toEqual([]);
});

test('no-op cannot be approved without reading original text', async () => {
    const send = scripted([contract, call('read_draft'), call('validate_patch'), finish]);
    await expect(runDocumentEditSession({ snapshot, instruction: 'Ensure limitations are mentioned.', send, maxSteps: 4 }))
        .rejects.toThrow(/did not complete/);
    expect(send.mock.calls.some(([messages]) => JSON.stringify(messages).includes('Read relevant original blocks'))).toBe(true);
});

test.each([
    review({ checks: [] }), review({ instructionSatisfied: false }), review({ issues: ['Missing support.'] }),
    review({ checks: [{ requirement: 0, satisfied: true, evidence: '' }] }), '{}', 'bad JSON',
])('incomplete or malformed review cannot produce a valid proposal', async (badReview) => {
    const send = scripted([contract, read, insert, call('read_draft'), call('validate_patch'), badReview, finish]);
    await expect(runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send, maxSteps: 6 })).rejects.toThrow(/did not complete/);
});

test('draft edits invalidate successful validation', async () => {
    const send = scripted([contract, read, insert, call('read_draft'), call('validate_patch'), review(),
        call('stage_patch', { operations: [{ kind: 'replace', blockId: 'draft-1', text: 'Different text.', reason: 'Revise.' }] }), finish]);
    await expect(runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send, maxSteps: 7 })).rejects.toThrow(/did not complete/);
});

test('reads only supplied source tools and passes successful source excerpts to review', async () => {
    const executeSource = jest.fn(async () => ({ ok: true, result: { text: 'Evidence excerpt.', offset: 12 } }));
    const send = scripted([call('file_read', { fileId: 'f', offset: 12 }), contract, read, insert,
        call('read_draft'), call('validate_patch'), review(), finish]);
    await runDocumentEditSession({ snapshot, instruction: 'Use the attached evidence.', sourceContext: 'attached f', send,
        sourceTools: [{ name: 'file_read', description: 'Read attached file', argsExample: {} }], executeSource });
    const request = send.mock.calls.find(([messages]) => messages[0].content.startsWith('Review a PROPOSED'))[0];
    expect(request[1].content).toContain('Evidence excerpt.');
    expect(executeSource).toHaveBeenCalledWith('file_read', { fileId: 'f', offset: 12 });
});

test('source failures and oversize evidence cannot become review evidence', async () => {
    const executeSource = jest.fn().mockResolvedValueOnce({ ok: false, error: 'Missing source' })
        .mockResolvedValueOnce({ ok: true, result: { text: 'x'.repeat(50000) } });
    const send = scripted([call('file_read', { fileId: 'f' }), call('file_read', { fileId: 'g' }), finish]);
    await expect(runDocumentEditSession({ snapshot, instruction: 'Use evidence.', send, maxSteps: 3,
        sourceTools: [{ name: 'file_read', description: 'read', argsExample: {} }], executeSource })).rejects.toThrow(/did not complete/);
});

test('review budget exhaustion prevents unbounded retries', async () => {
    const replies = [contract, read, insert, call('read_draft')];
    for (let i = 0; i < 3; i++) replies.push(call('validate_patch'), review({ satisfied: false }));
    replies.push(call('validate_patch'), finish);
    const send = scripted(replies);
    await expect(runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send, maxSteps: 9 })).rejects.toThrow(/did not complete/);
    expect(send.mock.calls.filter(([messages]) => messages[0].content.startsWith('Review a PROPOSED'))).toHaveLength(3);
});

test('aborts and transport failures never yield an applicable draft', async () => {
    const controller = new AbortController();
    const send = jest.fn(async () => { controller.abort(); return contract; });
    await expect(runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send, signal: controller.signal })).rejects.toMatchObject({ name: 'AbortError' });
    await expect(runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send: scripted([new Error('offline')]) })).rejects.toThrow('offline');
});

test.each(['', 'x'.repeat(48001)])('rejects invalid requests before sending', async (instruction) => {
    const send = jest.fn();
    await expect(runDocumentEditSession({ snapshot, instruction, send })).rejects.toThrow(/empty or too large/);
    expect(send).not.toHaveBeenCalled();
});
