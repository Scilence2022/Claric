/** Native paragraph moves. Verify the destination copy before removing its source. */
import { applyDocumentMove, resolveDocumentMove } from '../lib/document-move.js';
import { documentPartRoot } from '../lib/ooxml-text.js';
import { bodyElementFingerprint } from '../lib/ooxml-fingerprint.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const TYPES = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const PARAGRAPH_CHILDREN = new Set(['pPr', 'r', 'proofErr', 'bookmarkStart', 'bookmarkEnd']);
const RUN_CHILDREN = new Set(['rPr', 't', 'tab', 'br', 'cr', 'sym', 'noBreakHyphen', 'softHyphen', 'lastRenderedPageBreak']);
const UNSAFE_PROPERTIES = /^(?:sectPr|numPr|.*Change|ins|del|move.*)$/;
const RANGE_MARKER = /^(bookmark|commentRange|perm|moveFromRange|moveToRange|customXml\w+Range)(Start|End)$/;
const elements = (node, name) => Array.from(node.getElementsByTagNameNS(W, name));
const isParagraph = (node) => node.namespaceURI === W && node.localName === 'p';
const isClaricBookmark = (node) => (node.getAttributeNS(W, 'name') || '').startsWith('_claric_edit_');
const idOf = (node) => node.getAttributeNS(W, 'id');

function movableParagraph(paragraph) {
    if (!paragraph || !isParagraph(paragraph)) return false;
    const starts = elements(paragraph, 'bookmarkStart');
    const own = new Set(starts.filter(isClaricBookmark).map(idOf));
    return !starts.some((start) => !isClaricBookmark(start))
        && !elements(paragraph, 'bookmarkEnd').some((end) => !own.has(idOf(end)))
        && Array.from(paragraph.children).every((child) => child.namespaceURI === W && PARAGRAPH_CHILDREN.has(child.localName))
        && elements(paragraph, 'r').every((run) => Array.from(run.children).every((child) => child.namespaceURI === W
            && RUN_CHILDREN.has(child.localName) && (child.localName !== 'br' || ['textWrapping', null].includes(child.getAttributeNS(W, 'type')))))
        && Array.from(paragraph.getElementsByTagName('*')).every((child) => child.namespaceURI === W && !UNSAFE_PROPERTIES.test(child.localName));
}

export function canMoveParagraphXml(xml) {
    const root = documentPartRoot(xml);
    if (!root) return false;
    if (isParagraph(root)) return movableParagraph(root);
    const body = root.localName === 'body' && root.namespaceURI === W ? root : elements(root, 'body')[0];
    const paragraph = body && Array.from(body.children).find(isParagraph);
    return movableParagraph(paragraph);
}

function readMoveBody(xml) {
    const root = documentPartRoot(xml);
    const body = root && (root.localName === 'body' && root.namespaceURI === W ? root : elements(root, 'body')[0]);
    if (!body) throw new Error('Word could not verify the native document body for this move.');
    const active = new Set();
    const bookmarks = new Map();
    const own = new Set(elements(body, 'bookmarkStart').filter(isClaricBookmark).map(idOf));
    const nodes = Array.from(body.children).map((node, index) => {
        const beforeProtected = active.size > 0;
        let protectedRange = beforeProtected;
        for (const marker of elements(node, '*')) {
            const match = RANGE_MARKER.exec(marker.localName);
            if (!match) continue;
            const [, type, edge] = match;
            if (type === 'bookmark' && own.has(idOf(marker))) {
                let container = marker.parentElement;
                while (container?.namespaceURI === W && container.localName === 'ins') container = container.parentElement;
                if (edge === 'Start' && isParagraph(node) && container === node) bookmarks.set(marker.getAttributeNS(W, 'name'), index);
                continue;
            }
            const id = `${type}:${idOf(marker)}`;
            if (edge === 'Start') active.add(id);
            else active.delete(id);
            protectedRange = true;
        }
        return { node, signature: bodyElementFingerprint(node), beforeProtected, afterProtected: active.size > 0, protectedRange };
    });
    return { root, nodes, bookmarks };
}

function movePackage(root, source) {
    const serializer = new XMLSerializer();
    const paragraphs = source.map(({ node }) => {
        const copy = /** @type {Element} */ (node.cloneNode(true));
        for (const name of ['bookmarkStart', 'bookmarkEnd', 'proofErr', 'lastRenderedPageBreak']) {
            for (const el of elements(copy, name)) el.remove();
        }
        for (const attr of Array.from(copy.attributes)) {
            if (['paraId', 'textId'].includes(attr.localName)) copy.removeAttributeNode(attr);
        }
        return serializer.serializeToString(copy);
    }).join('');
    const stylePart = Array.from(root.ownerDocument.getElementsByTagNameNS(PKG, 'part'))
        .find((part) => part.getAttributeNS(PKG, 'name') === '/word/styles.xml');
    return `<pkg:package xmlns:pkg="${PKG}">`
        + `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">`
        + `<pkg:xmlData><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>`
        + `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${TYPES}.document.main+xml">`
        + `<pkg:xmlData><w:document xmlns:w="${W}"><w:body>${paragraphs}</w:body></w:document></pkg:xmlData></pkg:part>`
        + (stylePart ? `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">`
            + `<pkg:xmlData><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE_REL}/styles" Target="styles.xml"/></Relationships></pkg:xmlData></pkg:part>`
            + serializer.serializeToString(stylePart) : '') + '</pkg:package>';
}

function requireMoveSafety(source, target, side) {
    if (source.some((item) => item.protectedRange || !movableParagraph(item.node))) {
        throw new Error('Move source contains a protected Word object, range, section break, list or unresolved revision.');
    }
    if (!isParagraph(target.node) || (side === 'after' ? target.afterProtected : target.beforeProtected)) {
        throw new Error('Move destination is not an independent body paragraph boundary.');
    }
}

/** Match source and destination Word bookmarks to direct body children, then preflight the entire plan. */
export function planNativeMoves(xml, anchors, moves, inserts) {
    const body = readMoveBody(xml);
    const labels = body.nodes.map((_node, index) => `body-${index}`);
    for (const [id, entry] of Object.entries(anchors)) {
        const index = body.bookmarks.get(entry.bookmark);
        if (index === undefined) continue;
        if (!labels[index].startsWith('body-')) throw new Error('Word could not locate a distinct body paragraph for this move.');
        labels[index] = id;
    }
    const byId = new Map(labels.map((label, index) => [label, body.nodes[index]]));
    const used = new Set();
    let order = labels.slice();
    for (const move of moves) {
        order = applyDocumentMove(order, move);
        if (move.blockIds.some((id) => used.has(id))) throw new Error('A paragraph can be moved once per draft.');
        requireMoveSafety(move.blockIds.map((id) => byId.get(id)), byId.get(move.targetId), move.side);
        move.blockIds.forEach((id) => used.add(id));
    }
    const content = order.filter((id) => byId.get(id).node.localName !== 'sectPr');
    for (const { afterId, beforeId } of inserts) {
        const left = afterId ? content.indexOf(afterId) : -1;
        const right = beforeId ? content.indexOf(beforeId) : content.length;
        if ((afterId && left < 0) || left + 1 !== right
            || (afterId && byId.get(afterId).afterProtected) || (beforeId && byId.get(beforeId).beforeProtected)) {
            throw new Error('The proposed transition insertion is not at a verified gap between body paragraphs.');
        }
    }
    return { root: body.root, nodes: body.nodes.map((item, index) => ({ ...item, key: labels[index] })) };
}

/** Under Track Changes Word retains a deleted paragraph and marks its paragraph end. */
function revisionProjection(item, kind) {
    if (!isParagraph(item.node)) return null;
    const copy = /** @type {Element} */ (item.node.cloneNode(true));
    const marks = elements(copy, kind);
    if (!marks.some((mark) => mark.parentElement?.localName === 'rPr' && mark.parentElement?.parentElement?.localName === 'pPr')) return null;
    if (Array.from(copy.children).some((child) => RUN_CHILDREN.has(child.localName) || child.localName === 'r')) return null;
    for (const mark of marks) {
        if (mark.parentElement?.localName === 'rPr') { mark.remove(); continue; }
        if (mark.parentElement !== copy) return null;
        if (kind === 'del') {
            for (const text of elements(mark, 'delText')) {
                const restored = copy.ownerDocument.createElementNS(W, 'w:t');
                restored.textContent = text.textContent;
                text.replaceWith(restored);
            }
        }
        mark.replaceWith(...Array.from(mark.childNodes));
    }
    return bodyElementFingerprint(copy);
}

function matchMoveBody(body, expected, kind, tracked) {
    if (body.nodes.length !== expected.length) throw new Error('Word read-back returned a different paragraph order or count.');
    return body.nodes.map((item, index) => {
        const baseline = expected[index];
        const signature = baseline.revised && tracked ? revisionProjection(item, kind) : item.signature;
        if (signature !== baseline.signature) throw new Error('Word read-back could not verify the moved content, formatting or tracked revisions.');
        return { ...item, key: baseline.key };
    });
}

/** Apply the preflighted operations in order and rebind original IDs to their new Word paragraphs. */
export async function applyNativeMoves(context, state, moves, targets, { tracked, check, onWrite, onApplied, createBookmark }) {
    const whole = (id) => targets.get(id).getRange(Word.RangeLocation.whole || 'Whole');
    for (const move of moves) {
        check();
        const { start, end, gap } = resolveDocumentMove(state.nodes.map((item) => item.key), move);
        const source = state.nodes.slice(start, end);
        const target = whole(move.targetId);
        if (typeof target.insertOoxml !== 'function' || typeof whole(move.blockIds[0]).delete !== 'function') {
            throw new Error('This Word host cannot move whole paragraphs.');
        }
        const payload = movePackage(state.root, source);
        const placeholderKeys = source.map((item) => `pending-${item.key}`);
        const expectedCopy = state.nodes.slice();
        expectedCopy.splice(gap, 0, ...source.map((item, index) => ({ ...item, key: placeholderKeys[index], revised: true })));
        const location = move.side === 'before' ? Word.InsertLocation.before : Word.InsertLocation.after;
        onWrite();
        const inserted = target.insertOoxml(payload, location);
        const paragraphs = inserted.paragraphs;
        paragraphs.load('items');
        const insertedXml = context.document.body.getOoxml();
        await context.sync();
        check();
        const afterCopy = readMoveBody(insertedXml.value);
        const copiedNodes = matchMoveBody(afterCopy, expectedCopy, 'ins', tracked);
        if (paragraphs.items.length !== source.length) throw new Error('Word could not address the inserted paragraphs. The source was preserved.');
        const copyBookmarks = paragraphs.items.map((paragraph) => {
            const name = createBookmark();
            paragraph.getRange(Word.RangeLocation.content).insertBookmark(name);
            return name;
        });
        const anchoredXml = context.document.body.getOoxml();
        await context.sync();
        check();
        const anchoredCopy = readMoveBody(anchoredXml.value);
        matchMoveBody(anchoredCopy, copiedNodes, 'ins', tracked);
        if (copyBookmarks.some((name, index) => anchoredCopy.bookmarks.get(name) !== gap + index)) {
            throw new Error('Word could not confirm the identities of the inserted paragraphs. The source was preserved.');
        }

        const keys = new Set(move.blockIds);
        const expectedFinal = copiedNodes.flatMap((item) => {
            if (keys.has(item.key)) return tracked ? [{ ...item, key: `deleted-${item.key}`, revised: true }] : [];
            const copied = placeholderKeys.indexOf(item.key);
            return [copied < 0 ? item : { ...item, key: move.blockIds[copied] }];
        });
        whole(move.blockIds[0]).expandTo(whole(move.blockIds.at(-1))).delete();
        const finalXml = context.document.body.getOoxml();
        await context.sync();
        check();
        const afterDelete = readMoveBody(finalXml.value);
        state.nodes = matchMoveBody(afterDelete, expectedFinal, 'del', tracked);
        state.root = afterDelete.root;
        move.blockIds.forEach((id, index) => targets.set(id, paragraphs.items[index]));
        onApplied(move.id);
    }
}
