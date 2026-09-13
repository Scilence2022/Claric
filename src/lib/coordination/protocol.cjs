const { Buffer } = require('buffer');
const { normalizeId } = require('./identity.cjs');
const PROTOCOL_VERSION = 1;
const PROTOCOL_V2 = 2;
const EVENT_TYPES = Object.freeze({
  PRESENCE: 'presence', METADATA: 'metadata',
  DOCUMENT_ANNOUNCE: 'document.announce', DOCUMENT_HEARTBEAT: 'document.heartbeat', DOCUMENT_LEFT: 'document.left',
  CONTEXT_REQUEST: 'context.request', CONTEXT_RESPONSE: 'context.response', CONTEXT_CANCEL: 'context.cancel',
  TASK_SUBMIT: 'task.submit', TASK_CLAIM: 'task.claim', TASK_PROGRESS: 'task.progress', TASK_SUCCEEDED: 'task.succeeded', TASK_FAILED: 'task.failed', TASK_CANCEL: 'task.cancel',
  PROPOSAL_ANNOUNCED: 'proposal.announced', PROPOSAL_UPDATED: 'proposal.updated', PROPOSAL_DECISION: 'proposal.decision', PROPOSAL_APPLIED: 'proposal.applied', PROPOSAL_CONFLICT: 'proposal.conflict',
  LEASE_ACQUIRE: 'lease.acquire', LEASE_RENEW: 'lease.renew', LEASE_RELEASE: 'lease.release',
  ARTIFACT_ANNOUNCED: 'artifact.announced', ARTIFACT_REMOVED: 'artifact.removed',
});
const MAX_EVENT_BYTES = 64 * 1024;
const MAX_PAYLOAD_BYTES = 48 * 1024;
const MAX_PAYLOAD_DEPTH = 4;
const MAX_TTL_MS = 86400000;
const SENSITIVE_KEY = /^(?:__proto__|constructor|prototype|password|secret|token|api[-_]?key|authorization|cookie|private[-_]?key|credential|body|full[-_]?text|document[-_]?text|plaintext|image[-_]?bytes|bytes|base64|range|ooxml|html)$/i;
const V2_EVENT_TYPES = new Set(Object.values(EVENT_TYPES).filter((type) => !['presence', 'metadata'].includes(type)));
const ALLOWED_FIELDS = new Set(['version', 'workspaceId', 'documentId', 'type', 'clientId', 'payload']);
const V2_FIELDS = new Set(['version', 'workspaceId', 'source', 'target', 'type', 'correlationId', 'idempotencyKey', 'ttlMs', 'createdAt', 'payload']);
const PARTY_FIELDS = ['workspaceId', 'documentId', 'instanceId'];
const PAYLOAD_SCHEMAS = Object.freeze({
  'document.announce': { title: 'string?', revision: 'id?', capabilities: 'ids?' },
  'document.heartbeat': { revision: 'id?' }, 'document.left': {},
  'context.request': { requestId: 'id', scope: 'id', artifactIds: 'ids?' },
  'context.response': { requestId: 'id', artifactIds: 'ids?', summary: 'string?', snapshot: 'object?', error: 'object?' },
  'context.cancel': { requestId: 'id', reason: 'string?' },
  'task.submit': { taskId: 'id', taskType: 'id?', instruction: 'string?', artifactIds: 'ids?', proposalId: 'id?', reviewRequired: 'boolean?', graphId: 'id?', attemptId: 'id?', payload: 'object?' },
  'task.claim': { taskId: 'id', graphId: 'id?', attemptId: 'id?' },
  'task.progress': { taskId: 'id', progress: 'progress', message: 'string?' },
  'task.succeeded': { taskId: 'id', proposalId: 'id?', artifactIds: 'ids?' },
  'task.failed': { taskId: 'id', errorCode: 'id', message: 'string?' },
  'task.cancel': { taskId: 'id', reason: 'string?' },
  'proposal.announced': { proposalId: 'id', taskId: 'id', revision: 'integer', baseRevision: 'id', artifactIds: 'ids', kind: 'id?', scope: 'id?', title: 'string?', summary: 'string?' },
  'proposal.updated': { proposalId: 'id', revision: 'integer', baseRevision: 'id', artifactIds: 'ids', kind: 'id?', scope: 'id?', title: 'string?', summary: 'string?' },
  'proposal.decision': { proposalId: 'id', revision: 'integer', decision: 'decision', resourceId: 'resource?', fence: 'fence?' },
  'proposal.applied': { proposalId: 'id', revision: 'integer', documentRevision: 'id', resourceId: 'resource', fence: 'fence' },
  'proposal.conflict': { proposalId: 'id', revision: 'integer', reason: 'id' },
  'lease.acquire': { resourceId: 'resource', durationMs: 'duration' },
  'lease.renew': { resourceId: 'resource', durationMs: 'duration', fence: 'fence' },
  'lease.release': { resourceId: 'resource', fence: 'fence' },
  'artifact.announced': { artifactId: 'id', mediaType: 'string', digest: 'id', size: 'integer', summary: 'string?' },
  'artifact.removed': { artifactId: 'id' },
});
function plainObject(value) { return value !== null && typeof value === 'object' && [Object.prototype, null].includes(Object.getPrototypeOf(value)); }
function validatePayload(value, depth = 0, strict = true) {
  if (depth > MAX_PAYLOAD_DEPTH) return false;
  if (value === null || typeof value === 'boolean') return true;
  if (typeof value === 'string') return strict ? !/data:[^,]*;base64,/i.test(value) : true;
  if (typeof value === 'number') return Number.isFinite(value);
  if (Array.isArray(value)) return value.length <= 100 && value.every((child) => validatePayload(child, depth + 1, strict));
  if (!plainObject(value)) return false;
  return Object.entries(value).every(([key, child]) => (!strict || !SENSITIVE_KEY.test(key)) && validatePayload(child, depth + 1, strict));
}
function normalizeParty(value, field) {
  if (!plainObject(value) || Object.keys(value).some((key) => !PARTY_FIELDS.includes(key))) throw new Error(`Invalid ${field}`);
  return Object.fromEntries(PARTY_FIELDS.map((key) => [key, normalizeId(value[key], `${field}.${key}`)]));
}
function validateSchema(payload, schema) {
  for (const key of Object.keys(payload)) if (!Object.hasOwn(schema, key)) throw new Error(`Unknown payload field: ${key}`);
  for (const [key, rule] of Object.entries(schema)) {
    const value = payload[key];
    if (value === undefined && rule.endsWith('?')) continue;
    const type = rule.replace('?', '');
    if (type === 'id') { normalizeId(value, key); continue; }
    const valid = type === 'string' ? typeof value === 'string' && value.length <= 2048
      : type === 'boolean' ? typeof value === 'boolean'
        : type === 'object' ? plainObject(value)
        : type === 'ids' ? Array.isArray(value) && value.length <= 100 && value.every((id) => normalizeId(id, key))
        : type === 'decision' ? ['accepted', 'rejected'].includes(value)
          : type === 'resource' ? value === 'write'
            : type === 'fence' ? Number.isSafeInteger(value) && value > 0
              : type === 'progress' ? Number.isFinite(value) && value >= 0 && value <= 1
                : type === 'duration' ? Number.isInteger(value) && value >= 1000 && value <= 60000
                  : Number.isSafeInteger(value) && value >= 0;
    if (!valid) throw new Error(`Invalid payload ${key}`);
  }
}
function validateV2Envelope(input, now = Date.now()) {
  if (!plainObject(input)) throw new Error('Invalid envelope');
  for (const key of Object.keys(input)) if (!V2_FIELDS.has(key)) throw new Error(`Unknown envelope field: ${key}`);
  if (input.version !== PROTOCOL_V2 || !V2_EVENT_TYPES.has(input.type)) throw new Error('Unsupported envelope');
  const source = normalizeParty(input.source, 'source');
  const target = normalizeParty(input.target, 'target');
  const workspaceId = normalizeId(input.workspaceId, 'workspaceId');
  if (source.workspaceId !== workspaceId || target.workspaceId !== workspaceId) throw new Error('Workspace mismatch');
  normalizeId(input.correlationId, 'correlationId'); normalizeId(input.idempotencyKey, 'idempotencyKey');
  if (!Number.isInteger(input.ttlMs) || input.ttlMs < 1000 || input.ttlMs > MAX_TTL_MS) throw new Error('Invalid TTL');
  if (!Number.isSafeInteger(input.createdAt) || input.createdAt < 0 || input.createdAt > now + 30000 || input.createdAt + input.ttlMs <= now) throw new Error('Expired or invalid createdAt');
  if (!plainObject(input.payload) || !validatePayload(input.payload)) throw new Error('Invalid event payload');
  if (Buffer.byteLength(JSON.stringify(input.payload)) > MAX_PAYLOAD_BYTES || Buffer.byteLength(JSON.stringify(input)) > MAX_EVENT_BYTES) throw new Error('Event too large');
  validateSchema(input.payload, PAYLOAD_SCHEMAS[input.type]);
  return JSON.parse(JSON.stringify({ ...input, source, target }));
}
function validateEvent(input, workspaceId) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) throw new Error('Invalid event');
  for (const key of Object.keys(input)) if (!ALLOWED_FIELDS.has(key)) throw new Error(`Unknown event field: ${key}`);
  if (input.version !== PROTOCOL_VERSION) throw new Error('Unsupported event version');
  const eventWorkspace = normalizeId(input.workspaceId, 'workspaceId');
  if (eventWorkspace !== workspaceId) throw new Error('Workspace mismatch');
  const type = input.type;
  if (!['presence', 'metadata'].includes(type)) throw new Error('Invalid event type');
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
module.exports = { PROTOCOL_VERSION, PROTOCOL_V2, EVENT_TYPES, V2_EVENT_TYPES, MAX_EVENT_BYTES, MAX_PAYLOAD_BYTES, validateEvent, validateV2Envelope, createEnvelope };
