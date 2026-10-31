/** Shared client/server vocabulary. A peer restriction never overrides a master switch. */
const ACCESS_OPERATIONS = ['readContext', 'requestEdits'];
function normalizeDocumentAccess(value = {}) {
  const object = (v) => v && typeof v === 'object' && !Array.isArray(v);
  if (!object(value) || Object.keys(value).some((key) => ![...ACCESS_OPERATIONS, 'peers'].includes(key))) throw new Error('Invalid document permissions');
  const result = { readContext: true, requestEdits: true, peers: [] };
  for (const key of ACCESS_OPERATIONS) if (value[key] !== undefined) {
    if (typeof value[key] !== 'boolean') throw new Error('Invalid document permission');
    result[key] = value[key];
  }
  if (value.peers !== undefined) {
    if (!Array.isArray(value.peers) || value.peers.length > 100) throw new Error('Too many document permission rules');
    const seen = new Set();
    result.peers = value.peers.map((peer) => {
      if (!object(peer) || Object.keys(peer).some((key) => !['documentId', ...ACCESS_OPERATIONS].includes(key))
        || !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/.test(peer.documentId) || seen.has(peer.documentId)) throw new Error('Invalid peer permission rule');
      seen.add(peer.documentId);
      const item = { documentId: peer.documentId };
      for (const key of ACCESS_OPERATIONS) if (peer[key] !== undefined) {
        if (typeof peer[key] !== 'boolean') throw new Error('Invalid peer permission');
        item[key] = peer[key];
      }
      return item;
    });
  }
  return result;
}
function allowsDocumentAccess(value, documentId, operation) {
  if (!ACCESS_OPERATIONS.includes(operation)) return false;
  try {
    const access = normalizeDocumentAccess(value);
    return access[operation] && access.peers.find((peer) => peer.documentId === documentId)?.[operation] !== false;
  } catch { return false; }
}
module.exports = { ACCESS_OPERATIONS, normalizeDocumentAccess, allowsDocumentAccess };
