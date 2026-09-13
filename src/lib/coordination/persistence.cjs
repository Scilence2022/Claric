const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { Buffer } = require('buffer');
const { normalizeId } = require('./identity.cjs');
const processInfo = globalThis.process || {};

const PROJECT_ROOT = path.resolve(path.dirname(require.resolve('./persistence.cjs')), '../../..');
const STATE_BASENAME = 'coordination-checkpoint.json';
const MAX_BYTES = 16 * 1024 * 1024;
const MAX_RECORDS = 10000;
const TERMINAL = { tasks: ['succeeded', 'failed', 'cancelled'], proposals: ['applied', 'rejected', 'conflict'] };
const STATES = { tasks: [...TERMINAL.tasks, 'submitted', 'claimed', 'running', 'interrupted', 'unknown'], proposals: [...TERMINAL.proposals, 'announced', 'accepted', 'interrupted', 'unknown'] };
const ID_FIELDS = ['taskId', 'proposalId', 'scope', 'correlationId', 'attemptId', 'graphId', 'baseRevision', 'documentRevision'];
const NUMBER_FIELDS = ['revision', 'createdAt', 'updatedAt', 'expiresAt', 'fence'];
const FIELDS = ['workspaceId', 'source', 'target', 'state', 'epoch', 'digest', 'previousState', ...ID_FIELDS, ...NUMBER_FIELDS];
function persistenceError(message) { return Object.assign(new Error(message), { status: 503, code: 'CHECKPOINT_UNAVAILABLE' }); }
function object(value) { return !!value && typeof value === 'object' && !Array.isArray(value); }
function exact(value, fields) { if (!object(value) || Object.keys(value).some((field) => !fields.includes(field))) throw new Error('Invalid checkpoint fields'); }
function integer(value) { if (!Number.isSafeInteger(value) || value < 0) throw new Error('Invalid checkpoint integer'); }
function digest(value) { return crypto.createHash('sha256').update(value).digest('hex'); }
function safeStatePath(filePath) {
  if (typeof filePath !== 'string' || !filePath) throw persistenceError('Configure an explicit coordination checkpoint path');
  const resolved = path.resolve(PROJECT_ROOT, filePath);
  const relative = path.relative(PROJECT_ROOT, resolved);
  const parts = relative.split(path.sep);
  if (relative.startsWith('..') || path.isAbsolute(relative) || parts.length < 2 || path.basename(resolved) !== STATE_BASENAME || parts.some((part) => part.startsWith('.'))) throw persistenceError(`Checkpoint must be ${STATE_BASENAME} in a non-hidden data directory inside the project root`);
  let current = PROJECT_ROOT;
  for (const part of parts.slice(0, -1)) {
    current = path.join(current, part);
    const info = fs.lstatSync(current);
    if (!info.isDirectory() || info.isSymbolicLink()) throw persistenceError('Checkpoint directory must not use symlinks');
  }
  const directory = fs.statSync(path.dirname(resolved));
  if ((directory.mode & 0o077) !== 0 || directory.uid !== processInfo.getuid?.()) throw persistenceError('Checkpoint data directory must be owned by this user and mode 0700');
  return resolved;
}
function readCheckpoint(filePath) {
  safeStatePath(filePath);
  let fd;
  try {
    fd = fs.openSync(filePath, fs.constants.O_RDONLY | fs.constants.O_NOFOLLOW | fs.constants.O_NONBLOCK);
    const info = fs.fstatSync(fd);
    if (!info.isFile() || info.nlink !== 1 || info.size > MAX_BYTES || (info.mode & 0o777) !== 0o600 || info.uid !== processInfo.getuid?.()) throw persistenceError('Checkpoint must be a private, bounded regular file');
    return fs.readFileSync(fd, 'utf8');
  } catch (error) { if (error.code === 'ENOENT') return null; throw error; }
  finally { if (fd !== undefined) fs.closeSync(fd); }
}
function validateRecord(item, group) {
  exact(item, FIELDS);
  for (const field of ['workspaceId', 'epoch', 'correlationId', 'taskId', ...(group === 'proposals' ? ['proposalId', 'baseRevision'] : [])]) normalizeId(item[field], field);
  for (const field of ID_FIELDS) if (item[field] !== undefined) normalizeId(item[field], field);
  for (const field of NUMBER_FIELDS) if (item[field] !== undefined) integer(item[field]);
  integer(item.createdAt); integer(item.expiresAt);
  if (group === 'proposals') integer(item.revision);
  for (const party of ['source', 'target']) {
    exact(item[party], ['workspaceId', 'documentId', 'instanceId']);
    for (const field of ['workspaceId', 'documentId', 'instanceId']) normalizeId(item[party][field], field);
    if (item[party].workspaceId !== item.workspaceId) throw new Error('Checkpoint workspace mismatch');
  }
  if (!STATES[group].includes(item.state) || item.previousState !== undefined && !STATES[group].includes(item.previousState) || !/^[a-f0-9]{64}$/.test(item.digest)) throw new Error('Invalid checkpoint state or digest');
}
function validateCheckpoint(value, maxRecords) {
  exact(value, ['version', 'epoch', 'savedAt', 'highWater', 'tasks', 'proposals']);
  if (value.version !== 1) throw new Error('Unsupported checkpoint version');
  normalizeId(value.epoch, 'epoch'); integer(value.savedAt);
  exact(value.highWater, ['fence']); integer(value.highWater.fence);
  if (!Array.isArray(value.tasks) || !Array.isArray(value.proposals) || value.tasks.length + value.proposals.length > maxRecords) throw new Error('Checkpoint record capacity exceeded');
  for (const group of ['tasks', 'proposals']) {
    const seen = new Set();
    for (const item of value[group]) {
      validateRecord(item, group);
      const id = JSON.stringify([item.workspaceId, item[group === 'tasks' ? 'taskId' : 'proposalId']]);
      if (seen.has(id)) throw new Error('Duplicate checkpoint record');
      seen.add(id);
    }
  }
  return value;
}
function metadata(item, epoch) {
  const result = { epoch, digest: digest(JSON.stringify(item)) };
  for (const field of [...ID_FIELDS, ...NUMBER_FIELDS, 'workspaceId', 'state']) if (item[field] !== undefined) result[field] = item[field];
  for (const party of ['source', 'target']) result[party] = Object.fromEntries(['workspaceId', 'documentId', 'instanceId'].map((field) => [field, item[party][field]]));
  return result;
}
function recover(records, group) {
  return records.map((item) => {
    if (TERMINAL[group].includes(item.state) || ['interrupted', 'unknown'].includes(item.state)) return item;
    return { ...item, previousState: item.state, state: item.state === 'accepted' ? 'unknown' : 'interrupted' };
  });
}
class CoordinationPersistence {
  constructor(filePath, { maxRecords = MAX_RECORDS } = {}) {
    this.maxRecords = Math.min(maxRecords, MAX_RECORDS);
    try {
      this.filePath = safeStatePath(filePath);
      const raw = readCheckpoint(this.filePath);
      const saved = raw === null ? null : validateCheckpoint(JSON.parse(raw), this.maxRecords);
      this.expectedDigest = raw === null ? null : digest(raw);
      this.highWater = saved?.highWater.fence || 0;
      this.recovery = { tasks: recover(saved?.tasks || [], 'tasks'), proposals: recover(saved?.proposals || [], 'proposals') };
    } catch { throw persistenceError('Checkpoint startup refused: unsafe path, permissions, or invalid checkpoint; existing data was not modified'); }
  }
  checkpoint(state) {
    let temporary;
    let fd;
    let directoryFd;
    let renamed = false;
    try {
      const value = { version: 1, epoch: state.cursors.epoch, savedAt: state.clock(), highWater: { fence: state.leases.nextFence } };
      for (const group of ['tasks', 'proposals']) value[group] = [...state.recovery[group], ...[...state[group].records.values()].map((item) => metadata(item, state.cursors.epoch))];
      const data = JSON.stringify(validateCheckpoint(value, this.maxRecords));
      if (Buffer.byteLength(data) > MAX_BYTES) throw new Error('Checkpoint byte capacity exceeded');
      const previous = readCheckpoint(this.filePath);
      if ((previous === null ? null : digest(previous)) !== this.expectedDigest) throw new Error('Checkpoint changed externally');
      temporary = path.join(path.dirname(this.filePath), `${STATE_BASENAME}.${crypto.randomBytes(16).toString('hex')}.tmp`);
      fd = fs.openSync(temporary, fs.constants.O_WRONLY | fs.constants.O_CREAT | fs.constants.O_EXCL | fs.constants.O_NOFOLLOW, 0o600);
      fs.fchmodSync(fd, 0o600);
      fs.writeFileSync(fd, data, 'utf8');
      fs.fsyncSync(fd); fs.closeSync(fd); fd = undefined;
      directoryFd = fs.openSync(path.dirname(this.filePath), fs.constants.O_RDONLY | fs.constants.O_NOFOLLOW);
      safeStatePath(this.filePath);
      fs.renameSync(temporary, this.filePath); renamed = true;
      fs.fsyncSync(directoryFd);
      this.expectedDigest = digest(data);
    } catch {
      throw Object.assign(persistenceError('Checkpoint failed; coordination writes are blocked until restart and local application outcome must be reconciled'), { commitUnknown: renamed });
    } finally {
      if (fd !== undefined) { try { fs.closeSync(fd); } catch (error) { void error; } }
      if (directoryFd !== undefined) { try { fs.closeSync(directoryFd); } catch (error) { void error; } }
      if (temporary && !renamed) { try { fs.unlinkSync(temporary); } catch (error) { void error; } }
    }
  }
}
module.exports = { CoordinationPersistence, STATE_BASENAME, PROJECT_ROOT };
