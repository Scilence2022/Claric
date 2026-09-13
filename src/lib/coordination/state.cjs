const crypto = require('crypto');
const { normalizeId } = require('./identity.cjs');
const { validateV2Envelope } = require('./protocol.cjs');
const { CoordinationPersistence } = require('./persistence.cjs');

function clone(value) { return value === undefined ? value : JSON.parse(JSON.stringify(value)); }
function key(...parts) { return parts.join('\u0000'); }
function error(message, status = 409) { return Object.assign(new Error(message), { status }); }
function same(a, b) { return !!a && !!b && ['workspaceId', 'documentId', 'instanceId'].every((field) => a[field] === b[field]); }
function requireOwner(actual, expected) { if (!same(actual, expected)) throw error('Session does not own this operation', 403); }
function canonical(value) {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`;
  if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((name) => `${JSON.stringify(name)}:${canonical(value[name])}`).join(',')}}`;
  return JSON.stringify(value);
}

class SessionStore {
  constructor({ clock = Date.now, maxRecords = 10000, sessionTtlMs = 86400000 } = {}) {
    this.clock = clock; this.maxRecords = maxRecords; this.sessionTtlMs = sessionTtlMs; this.sessions = new Map();
  }
  purge() { for (const [id, session] of this.sessions) if (session.expiresAt <= this.clock()) this.sessions.delete(id); }
  register(input, binding) {
    if (!input || typeof input !== 'object' || Array.isArray(input) || Object.keys(input).some((field) => !['workspaceId', 'documentId'].includes(field))) throw error('Invalid registration', 400);
    const workspaceId = normalizeId(input.workspaceId, 'workspaceId'); const documentId = normalizeId(input.documentId, 'documentId');
    this.purge(); if (this.sessions.size >= this.maxRecords) throw error('Session capacity reached', 503);
    const credential = crypto.randomBytes(32).toString('base64url');
    const identity = { workspaceId, documentId, instanceId: crypto.randomUUID() };
    const expiresAt = this.clock() + this.sessionTtlMs;
    this.sessions.set(this.digest(credential), { identity, binding, expiresAt });
    return { credential, identity, expiresAt };
  }
  digest(value) { return crypto.createHash('sha256').update(value).digest('hex'); }
  authenticate(credential, binding) {
    this.purge(); const session = typeof credential === 'string' && this.sessions.get(this.digest(credential));
    if (!session || session.binding !== binding) throw error('Invalid or expired instance session', 401);
    return clone(session.identity);
  }
  revoke(identity) { for (const [id, session] of this.sessions) if (same(identity, session.identity)) this.sessions.delete(id); }
  has(identity) { this.purge(); return [...this.sessions.values()].some((session) => same(identity, session.identity)); }
}

class RecordStore {
  constructor({ clock = Date.now, maxRecords = 10000 } = {}) { this.clock = clock; this.maxRecords = maxRecords; this.records = new Map(); }
  purge() { for (const [id, record] of this.records) if (record.expiresAt <= this.clock()) this.records.delete(id); }
  get(workspaceId, id) { this.purge(); return clone(this.records.get(key(workspaceId, id)) || null); }
  put(workspaceId, id, record) {
    this.purge(); const recordKey = key(workspaceId, id);
    if (!this.records.has(recordKey) && this.records.size >= this.maxRecords) throw error('State capacity reached', 503);
    this.records.set(recordKey, clone(record)); return clone(record);
  }
  remove(workspaceId, id) { this.records.delete(key(workspaceId, id)); }
  list(workspaceId, identity) {
    normalizeId(workspaceId, 'workspaceId'); this.purge();
    return [...this.records.values()].filter((item) => item.workspaceId === workspaceId && (!identity || same(item.source, identity) || same(item.target, identity))).map(clone);
  }
  require(envelope, id) {
    const item = this.get(envelope.workspaceId, id);
    if (!item) throw error('Record not found or expired', 404);
    if (item.correlationId !== envelope.correlationId) throw error('Correlation mismatch');
    return item;
  }
  create(envelope, id, state) {
    if (this.get(envelope.workspaceId, id)) throw error('Record already exists');
    return this.put(envelope.workspaceId, id, { ...clone(envelope.payload), workspaceId: envelope.workspaceId, source: envelope.source, target: envelope.target, correlationId: envelope.correlationId, state, createdAt: this.clock(), expiresAt: envelope.createdAt + envelope.ttlMs });
  }
  route(envelope, item, actor) {
    requireOwner(envelope.source, actor);
    requireOwner(envelope.target, same(actor, item.source) ? item.target : item.source);
  }
}

class DocumentRegistry extends RecordStore {
  constructor(options = {}) { super(options); this.presenceTtlMs = options.presenceTtlMs || 60000; }
  list(workspaceId) {
    normalizeId(workspaceId, 'workspaceId'); this.purge();
    return [...this.records.values()]
      .filter((item) => item.workspaceId === workspaceId && item.expiresAt > this.clock())
      .map(clone);
  }
  apply(envelope) {
    const identity = envelope.source; requireOwner(envelope.target, identity);
    if (envelope.type === 'document.left') { this.remove(identity.workspaceId, identity.instanceId); return { ...identity, active: false }; }
    const old = this.get(identity.workspaceId, identity.instanceId);
    if (envelope.type === 'document.heartbeat' && !old) throw error('Announce document before heartbeat', 404);
    return this.put(identity.workspaceId, identity.instanceId, { ...(old || {}), ...identity, ...envelope.payload, lastSeenAt: this.clock(), expiresAt: Math.min(this.clock() + this.presenceTtlMs, envelope.createdAt + envelope.ttlMs) });
  }
}

class RequestStore extends RecordStore {
  apply(envelope) {
    const { requestId } = envelope.payload;
    if (envelope.type === 'context.request') return this.create(envelope, requestId, 'pending');
    const item = this.require(envelope, requestId);
    this.route(envelope, item, envelope.type === 'context.cancel' ? item.source : item.target);
    if (item.state !== 'pending') throw error('Request is already terminal');
    return this.put(envelope.workspaceId, requestId, { ...item, result: envelope.payload, state: envelope.type === 'context.cancel' ? 'cancelled' : 'responded', updatedAt: this.clock() });
  }
}

class TaskStore extends RecordStore {
  apply(envelope) {
    const { taskId } = envelope.payload;
    if (envelope.type === 'task.submit') return this.create(envelope, taskId, 'submitted');
    const item = this.require(envelope, taskId);
    this.route(envelope, item, envelope.type === 'task.cancel' ? item.source : item.target);
    if (['succeeded', 'failed', 'cancelled'].includes(item.state)) throw error('Task is already terminal');
    let state;
    if (envelope.type === 'task.claim') {
      if (item.state !== 'submitted') throw error('Task already claimed');
      state = 'claimed'; item.claimedBy = envelope.source;
    } else if (envelope.type === 'task.cancel') state = 'cancelled';
    else {
      if (!['claimed', 'running'].includes(item.state)) throw error('Claim task before updating');
      requireOwner(envelope.source, item.claimedBy);
      state = { 'task.progress': 'running', 'task.succeeded': 'succeeded', 'task.failed': 'failed' }[envelope.type];
      if (envelope.type === 'task.progress' && envelope.payload.progress < (item.progress || 0)) throw error('Progress cannot decrease');
    }
    return this.put(envelope.workspaceId, taskId, { ...item, ...envelope.payload, state, updatedAt: this.clock() });
  }
}

class LeaseStore extends RecordStore {
  constructor(options = {}) { super(options); this.nextFence = 0; }
  apply(envelope) {
    const p = envelope.payload; const identity = envelope.source;
    requireOwner(envelope.target, identity);
    if (p.resourceId !== 'write') throw error('Lease resource must be write');
    const resourceId = key(identity.documentId, p.resourceId);
    const old = this.get(identity.workspaceId, resourceId);
    if (envelope.type === 'lease.acquire') {
      if (old && old.expiresAt > this.clock()) throw error('Lease unavailable');
      return this.put(identity.workspaceId, resourceId, { workspaceId: identity.workspaceId, documentId: identity.documentId, resourceId: p.resourceId, source: identity, fence: ++this.nextFence, expiresAt: Math.min(this.clock() + p.durationMs, envelope.createdAt + envelope.ttlMs) });
    }
    this.assert(identity, p.resourceId, p.fence);
    if (envelope.type === 'lease.release') { this.remove(identity.workspaceId, resourceId); return { ...old, released: true }; }
    return this.put(identity.workspaceId, resourceId, { ...old, expiresAt: Math.min(this.clock() + p.durationMs, envelope.createdAt + envelope.ttlMs) });
  }
  assert(identity, resourceId, fence) {
    if (resourceId !== 'write' || !Number.isSafeInteger(fence) || fence <= 0) throw error('Invalid lease fence or resource');
    const lease = this.get(identity.workspaceId, key(identity.documentId, resourceId));
    if (!lease || lease.fence !== fence || lease.expiresAt <= this.clock()) throw error('Lease missing, expired, or stale');
    requireOwner(identity, lease.source); return lease;
  }
  revoke(identity) { for (const [id, lease] of this.records) if (same(lease.source, identity)) this.records.delete(id); }
}

class ProposalStore extends RecordStore {
  apply(envelope, leases, tasks) {
    const p = envelope.payload;
    if (envelope.type === 'proposal.announced') {
      const task = tasks?.get(envelope.workspaceId, p.taskId);
      if (!task || task.state !== 'claimed') throw error('Proposal requires a claimed task', 409);
      requireOwner(envelope.source, task.target);
      requireOwner(envelope.source, task.claimedBy);
      requireOwner(envelope.target, task.source);
      if (task.correlationId !== envelope.correlationId) throw error('Correlation mismatch');
      if (envelope.createdAt + envelope.ttlMs > task.expiresAt) throw error('Proposal cannot outlive task');
      return this.create({ ...envelope, source: task.source, target: task.target }, p.proposalId, 'announced');
    }
    const item = this.require(envelope, p.proposalId);
    requireOwner(envelope.source, item.target);
    requireOwner(envelope.target, item.source);
    const task = tasks.require(envelope, item.taskId);
    if (!['claimed', 'running'].includes(task.state)) throw error('Task is not active');
    if (['applied', 'rejected', 'conflict'].includes(item.state)) throw error('Proposal is already terminal');
    if (envelope.type === 'proposal.updated') {
      if (p.revision !== item.revision + 1) throw error('Proposal revision must increase by one');
      return this.put(envelope.workspaceId, p.proposalId, { ...item, ...p, state: 'announced', decision: null, updatedAt: this.clock() });
    }
    if (p.revision !== item.revision) throw error('Stale proposal revision');
    let state;
    if (envelope.type === 'proposal.decision') {
      if (item.state !== 'announced') throw error('Proposal cannot be decided in current state');
      if (p.decision === 'accepted') leases.assert(envelope.source, p.resourceId, p.fence);
      state = p.decision;
    } else if (envelope.type === 'proposal.applied') {
      if (item.state !== 'accepted') throw error('Accept proposal before applying');
      if (p.resourceId !== 'write') throw error('Lease resource must be write');
      leases.assert(envelope.source, p.resourceId, p.fence); state = 'applied';
    } else state = 'conflict';
    return this.put(envelope.workspaceId, p.proposalId, { ...item, ...p, state, updatedAt: this.clock() });
  }
}

class ArtifactStore extends RecordStore {
  apply(envelope) {
    const p = envelope.payload;
    if (envelope.type === 'artifact.announced') return this.create(envelope, p.artifactId, 'available');
    const item = this.require(envelope, p.artifactId); this.route(envelope, item, item.source);
    this.remove(envelope.workspaceId, p.artifactId); return { artifactId: p.artifactId, state: 'removed' };
  }
}

class SnapshotCursorStore {
  constructor({ maxEvents = 1000, maxRooms = 1000, clock = Date.now } = {}) {
    if (!Number.isInteger(maxEvents) || maxEvents < 1 || !Number.isInteger(maxRooms) || maxRooms < 1) throw error('Invalid cursor capacity', 400);
    this.maxEvents = maxEvents; this.maxRooms = maxRooms; this.clock = clock; this.rooms = new Map(); this.epoch = crypto.randomUUID();
  }
  room(workspaceId) {
    normalizeId(workspaceId, 'workspaceId');
    if (!this.rooms.has(workspaceId)) {
      if (this.rooms.size >= this.maxRooms) throw error('Workspace capacity reached', 503);
      this.rooms.set(workspaceId, { sequence: 0, events: [] });
    }
    return this.rooms.get(workspaceId);
  }
  append(envelope, result) {
    const room = this.room(envelope.workspaceId); const event = { ...clone(envelope), result: clone(result), sequence: ++room.sequence, receivedAt: this.clock() };
    room.events.push(event); if (room.events.length > this.maxEvents) room.events.shift(); return clone(event);
  }
  events(workspaceId, after, identity, epoch) {
    const room = this.room(workspaceId); const first = room.events[0]?.sequence || room.sequence + 1;
    if (!Number.isSafeInteger(after) || after < 0) throw error('Invalid cursor', 400);
    if (epoch !== this.epoch || after > room.sequence || after < first - 1) throw error('Cursor expired; fetch a new snapshot', 410);
    return { epoch: this.epoch, cursor: room.sequence, events: room.events.filter((event) => event.sequence > after && event.createdAt + event.ttlMs > this.clock() && (same(event.source, identity) || same(event.target, identity))).map(clone) };
  }
}

class CoordinationState {
  constructor(options = {}) {
    this.clock = options.clock || Date.now; this.maxRecords = options.maxRecords || 10000;
    this.sessions = new SessionStore(options); this.documents = new DocumentRegistry(options);
    this.requests = new RequestStore(options); this.tasks = new TaskStore(options); this.proposals = new ProposalStore(options); this.leases = new LeaseStore(options); this.artifacts = new ArtifactStore(options);
    this.cursors = new SnapshotCursorStore(options); this.idempotency = new Map();
    this.persistence = options.persistence || (options.persistencePath ? new CoordinationPersistence(options.persistencePath, options) : null);
    this.recovery = this.persistence?.recovery || { tasks: [], proposals: [] };
    this.persistenceBlocked = false;
    if (this.persistence) this.leases.nextFence = this.persistence.highWater;
  }
  register(input, binding) { return this.sessions.register(input, binding); }
  authenticate(credential, binding) { return this.sessions.authenticate(credential, binding); }
  append(input, credential, binding) {
    if (this.persistenceBlocked) throw error('Coordination persistence unavailable; restart required', 503);
    const actor = this.authenticate(credential, binding);
    const envelope = validateV2Envelope(input, this.clock()); requireOwner(envelope.source, actor);
    if (!this.sessions.has(envelope.target)) throw error('Target session not registered', 404);
    for (const [id, entry] of this.idempotency) if (entry.expiresAt <= this.clock()) this.idempotency.delete(id);
    const id = key(actor.instanceId, envelope.idempotencyKey); const fingerprint = canonical(envelope); const old = this.idempotency.get(id);
    if (old) { if (old.fingerprint !== fingerprint) throw error('Idempotency key reused with different envelope'); return clone(old.event); }
    if (this.idempotency.size >= this.maxRecords) throw error('Idempotency capacity reached', 503);
    const stores = [this.documents, this.requests, this.tasks, this.proposals, this.leases, this.artifacts];
    const backup = stores.map((store) => ({ records: new Map(store.records), ...(store.nextFence === undefined ? {} : { nextFence: store.nextFence }) }));
    const cursorBackup = new Map([...this.cursors.rooms].map(([workspaceId, room]) => [workspaceId, { sequence: room.sequence, events: [...room.events] }]));
    try {
      this.cursors.room(envelope.workspaceId);
      const group = envelope.type.split('.')[0];
      const store = { document: this.documents, context: this.requests, task: this.tasks, proposal: this.proposals, lease: this.leases, artifact: this.artifacts }[group];
      const result = store.apply(envelope, this.leases, this.tasks);
      const event = this.cursors.append(envelope, result);
      this.idempotency.set(id, { fingerprint, event, expiresAt: envelope.createdAt + envelope.ttlMs });
      if (envelope.type === 'document.left') { this.leases.revoke(actor); this.sessions.revoke(actor); }
      if (this.persistence) this.persistence.checkpoint(this);
      return clone(event);
    } catch (caught) {
      for (let index = 0; index < stores.length; index += 1) { stores[index].records = backup[index].records; if (backup[index].nextFence !== undefined) stores[index].nextFence = backup[index].nextFence; }
      this.cursors.rooms = cursorBackup;
      this.idempotency.delete(id);
      if (caught.code === 'CHECKPOINT_UNAVAILABLE') this.persistenceBlocked = true;
      throw caught;
    }
  }
  snapshot(credential, binding) {
    const actor = this.authenticate(credential, binding); const room = this.cursors.room(actor.workspaceId);
    return { version: 2, storage: this.persistence ? 'local-checkpoint' : 'memory', epoch: this.cursors.epoch, workspaceId: actor.workspaceId, cursor: room.sequence, capturedAt: this.clock(), documents: this.documents.list(actor.workspaceId), requests: this.requests.list(actor.workspaceId, actor), tasks: this.tasks.list(actor.workspaceId, actor), proposals: this.proposals.list(actor.workspaceId, actor), leases: this.leases.list(actor.workspaceId, actor), artifacts: this.artifacts.list(actor.workspaceId, actor), recovery: { tasks: this.recovery.tasks.filter((item) => item.workspaceId === actor.workspaceId && (!item.source || same(item.source, actor) || same(item.target, actor))), proposals: this.recovery.proposals.filter((item) => item.workspaceId === actor.workspaceId && (!item.source || same(item.source, actor) || same(item.target, actor))) } };
  }
  events(credential, binding, after, epoch) { const actor = this.authenticate(credential, binding); return this.cursors.events(actor.workspaceId, after, actor, epoch); }
}
module.exports = { SessionStore, DocumentRegistry, RequestStore, TaskStore, ProposalStore, LeaseStore, ArtifactStore, SnapshotCursorStore, CoordinationState };
