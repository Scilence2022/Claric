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

test('a long document insertion retains early and late supporting facts after chat history eviction', async () => {
    const source = { id: 'large', blocks: Array.from({ length: 298 }, (_, index) => ({
        id: `p-${index + 1}`, text: `Background measurement ${index + 1}. ` + 'background '.repeat(140),
    })) };
    source.blocks[0] = { id: 'p-1', text: 'Overview', headingLevel: 1 };
    source.blocks[1].text = 'The project source is released under the BSD license for independent inspection.';
    source.blocks[269] = { id: 'p-270', text: 'Discussion', headingLevel: 1 };
    source.blocks[270].text = 'The mechanism enables reviewable proposals.';
    source.blocks[271].text = 'Operational limitations remain.';
    source.blocks[272].text = 'Independent reproduction requires the recorded test bundle.';
    source.blocks[273] = { id: 'p-274', text: 'Appendix', headingLevel: 1 };
    const replies = [];
    for (let start = 0; start < 192; start += 12) {
        replies.push(call('read_blocks', { ids: source.blocks.slice(start, start + 12).map((block) => block.id) }));
    }
    replies.push(call('read_blocks', { ids: ['p-271', 'p-272', 'p-273'] }),
        call('set_edit_contract', { goal: 'Discuss source access', requirements: ['Explain source access in a suitable section'] }),
        call('stage_patch', { operations: [{ kind: 'insert', afterId: 'p-271',
            paragraphs: ['BSD source access enables independent inspection; reproduction also requires the recorded test bundle.'],
            reason: 'Connect reviewability to the operational limitations.' }] }),
        call('read_draft'), call('validate_patch'), finish);
    const requests = [];
    const send = jest.fn(async (messages) => {
        requests.push(JSON.parse(JSON.stringify(messages)));
        if (messages[0].content.startsWith('Review a PROPOSED')) {
            const packet = JSON.parse(messages[1].content);
            expect(packet.sourceEvidence).toEqual([]);
            expect(packet.documentEvidence.map((block) => block.id)).toEqual(expect.arrayContaining(['p-2', 'p-273']));
            expect(JSON.stringify(packet.documentEvidence).length).toBeLessThanOrEqual(48000);
            return review({ summary: 'Original document excerpts support the focused insertion.' });
        }
        const reply = replies.shift();
        if (!reply) throw new Error('Unexpected agent request');
        return reply;
    });
    const onStep = jest.fn();
    const result = await runDocumentEditSession({ snapshot: source, instruction: 'Discuss the value of source access at a suitable location.', send, onStep });
    expect(result.status).toBe('staged');
    expect(result.patch.changes).toHaveLength(1);
    expect(result.patch.changes[0]).toMatchObject({ afterId: 'p-271', beforeId: 'p-272' });
    expect(result.toolLoop.steps).toBe(22);
    expect(requests.some((messages) => JSON.stringify(messages).includes('dropped to stay within the context budget'))).toBe(true);
    expect(JSON.parse(requests[0][1].content.split('\n\nBegin.')[0]).sectionIndex.blocks.map((block) => block.id))
        .toEqual(['p-1', 'p-270', 'p-274']);
    expect(onStep.mock.calls.some(([step]) => step.progress.phase === 'reviewing')).toBe(true);
    expect(onStep.mock.calls.some(([step]) => step.progress.maxSteps === 32)).toBe(true);
});

test('a reviewer can recover an omitted read source without an agent rewrite or an extra validation tool call', async () => {
    const source = { id: 's', blocks: [...snapshot.blocks,
        { id: 'fact', text: 'Historic supporting detail. ' + 'historic '.repeat(1400) },
        { id: 'large1', text: 'XXX discussion '.repeat(1350) },
        { id: 'large2', text: 'XXX discussion '.repeat(1350) },
    ] };
    const replies = [contract, read, call('read_blocks', { ids: ['fact'], limit: 24000 }),
        call('read_blocks', { ids: ['large1'], limit: 24000 }), call('read_blocks', { ids: ['large2'], limit: 24000 }),
        insert, call('read_draft'), call('validate_patch'), finish];
    let reviews = 0;
    const send = jest.fn(async (messages) => {
        if (messages[0].content.startsWith('Review a PROPOSED')) {
            const packet = JSON.parse(messages[1].content);
            reviews++;
            if (reviews === 1) {
                expect(packet.evidenceCoverage.omittedIds).toContain('fact');
                return review({ satisfied: false, issues: ['Need the historic source.'], neededEvidenceIds: ['fact'] });
            }
            expect(packet.documentEvidence.map((block) => block.id)).toContain('fact');
            return review();
        }
        return replies.shift();
    });
    const result = await runDocumentEditSession({ snapshot: source, instruction: 'Insert XXX.', send });
    expect(result.status).toBe('staged');
    expect(result.patch.revision).toBe(1);
    expect(reviews).toBe(2);
    expect(result.toolLoop.calls.filter((item) => item.tool === 'validate_patch')).toHaveLength(1);
});

test('reviewer requests for unread sources remain blocked instead of exposing unobserved document text', async () => {
    const source = { id: 's', blocks: [...snapshot.blocks, { id: 'unread', text: 'Unobserved fact.' }] };
    const send = scripted([contract, read, insert, call('read_draft'), call('validate_patch'),
        review({ satisfied: false, issues: ['Missing source.'], neededEvidenceIds: ['unread'] })]);
    await expect(runDocumentEditSession({ snapshot: source, instruction: 'Insert XXX.', send, maxSteps: 5 }))
        .rejects.toThrow(/Read the complete original block unread/);
    expect(send.mock.calls.filter(([messages]) => messages[0].content.startsWith('Review a PROPOSED'))).toHaveLength(1);
});

test('a passing final-step review completes the proposal without needing another model call', async () => {
    const send = scripted([contract, read, insert, call('read_draft'), call('validate_patch'), review()]);
    const result = await runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send, maxSteps: 5 });
    expect(result).toMatchObject({ status: 'staged', completedByValidation: true, summary: 'Suitable placement.' });
    expect(result.review.ok).toBe(true);
});

test('a failed or malformed later review cannot reuse an earlier passing review at the budget boundary', async () => {
    const send = scripted([contract, read, insert, call('read_draft'), call('validate_patch'), review(),
        call('validate_patch'), 'bad JSON']);
    await expect(runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send, maxSteps: 6 }))
        .rejects.toThrow(/did not complete/);
});

test('an unsuccessful operation after review still prevents completion at the budget boundary', async () => {
    const send = scripted([contract, read, insert, call('read_draft'), call('validate_patch'), review(),
        call('stage_patch', { operations: [{ kind: 'replace', blockId: 'unknown', text: 'Missing edit.', reason: 'Another required change.' }] })]);
    await expect(runDocumentEditSession({ snapshot, instruction: 'Insert XXX.', send, maxSteps: 6 }))
        .rejects.toThrow(/did not complete/);
});

test('pinning unread sources fails before spending model review budget', async () => {
    const source = { id: 's', blocks: [...snapshot.blocks, { id: 'unread', text: 'Unobserved fact.' }] };
    const send = scripted([contract, read, insert, call('read_draft'), call('validate_patch', { evidenceIds: ['unread'] })]);
    await expect(runDocumentEditSession({ snapshot: source, instruction: 'Insert XXX.', send, maxSteps: 5 }))
        .rejects.toThrow(/did not complete/);
    expect(send.mock.calls.some(([messages]) => messages[0].content.startsWith('Review a PROPOSED'))).toBe(false);
});
