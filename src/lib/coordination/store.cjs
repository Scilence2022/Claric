const { normalizeId } = require('./identity.cjs');
const { validateEvent } = require('./protocol.cjs');

function clone(value) {
  if (value === undefined) return value;
  return JSON.parse(JSON.stringify(value));
}

class CoordinationStore {
  constructor({ maxEvents = 100 } = {}) {
    if (!Number.isInteger(maxEvents) || maxEvents < 1) throw new Error('Invalid maxEvents');
    this.maxEvents = maxEvents;
    this.rooms = new Map();
  }
  room(workspaceId) {
    normalizeId(workspaceId, 'workspaceId');
    if (!this.rooms.has(workspaceId)) this.rooms.set(workspaceId, { sequence: 0, events: [], presence: {}, metadataByDocument: new Map() });
    return this.rooms.get(workspaceId);
  }
  append(input) {
    const workspaceId = normalizeId(input && input.workspaceId, 'workspaceId');
    const event = validateEvent(input, workspaceId);
    const room = this.room(workspaceId);
    const stored = { ...event, sequence: ++room.sequence, timestamp: new Date().toISOString(), payload: clone(event.payload) };
    room.events.push(stored);
    if (room.events.length > this.maxEvents) room.events.splice(0, room.events.length - this.maxEvents);
    if (event.type === 'presence') {
      if (event.payload.active === false) delete room.presence[event.clientId];
      else room.presence[event.clientId] = { clientId: event.clientId, documentId: event.documentId, ...clone(event.payload) };
    } else {
      const key = event.documentId || '*';
      const metadata = room.metadataByDocument.get(key) || {};
      Object.assign(metadata, clone(event.payload));
      room.metadataByDocument.set(key, metadata);
    }
    return clone(stored);
  }
  snapshot(workspaceId) {
    const room = this.room(workspaceId);
    const metadataByDocument = {};
    for (const [documentId, metadata] of room.metadataByDocument) metadataByDocument[documentId] = clone(metadata);
    return clone({ version: 1, workspaceId, sequence: room.sequence, presence: room.presence, metadataByDocument, events: room.events });
  }
}
module.exports = { CoordinationStore };
