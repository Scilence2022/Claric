const { Buffer } = require('buffer');
const { normalizeId } = require('./identity.cjs');
const PROTOCOL_VERSION = 1;
const EVENT_TYPES = Object.freeze({ PRESENCE: 'presence', METADATA: 'metadata' });
const MAX_EVENT_BYTES = 64 * 1024;
const ALLOWED_FIELDS = new Set(['version', 'workspaceId', 'documentId', 'type', 'clientId', 'payload']);
function validatePayload(value, depth = 0) {
  if (depth > 4 || value === null || typeof value !== 'object') return depth <= 4;
  if (Array.isArray(value)) return value.length <= 100 && value.every((child) => validatePayload(child, depth + 1));
  for (const [key, child] of Object.entries(value)) {
    if (key === '__proto__' || key === 'constructor' || key === 'prototype') return false;
    if (!validatePayload(child, depth + 1)) return false;
  }
  return true;
}
function validateEvent(input, workspaceId) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('Invalid event');
  for (const key of Object.keys(input)) if (!ALLOWED_FIELDS.has(key)) throw new Error(`Unknown event field: ${key}`);
  if (input.version !== PROTOCOL_VERSION) throw new Error('Unsupported event version');
  const eventWorkspace = normalizeId(input.workspaceId, 'workspaceId');
  if (eventWorkspace !== workspaceId) throw new Error('Workspace mismatch');
  const type = input.type;
  if (!Object.values(EVENT_TYPES).includes(type)) throw new Error('Invalid event type');
  const clientId = normalizeId(input.clientId, 'clientId');
  const documentId = input.documentId == null ? null : normalizeId(input.documentId, 'documentId');
  if (!input.payload || typeof input.payload !== 'object' || Array.isArray(input.payload) || !validatePayload(input.payload)) throw new Error('Invalid event payload');
  const encoded = JSON.stringify(input);
  if (Buffer.byteLength(encoded) > MAX_EVENT_BYTES) throw new Error('Event too large');
  return { version: PROTOCOL_VERSION, workspaceId: eventWorkspace, documentId, type, clientId, payload: input.payload };
}
function createEnvelope({ workspaceId, documentId, type, clientId, payload }) {
  return validateEvent({ version: PROTOCOL_VERSION, workspaceId, documentId, type, clientId, payload }, workspaceId);
}
module.exports = { PROTOCOL_VERSION, EVENT_TYPES, MAX_EVENT_BYTES, validateEvent, createEnvelope };
