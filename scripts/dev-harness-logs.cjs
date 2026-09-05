const { bounded, MAX_BYTES } = require('./dev-harness-store.cjs');

const LOG_PATH = 'logs/e2e-test-logs.json';

function createLogs(store) {
  function read() {
    const value = store.read(LOG_PATH, []);
    if (Array.isArray(value)) {
      if (value.some(entry => !entry || typeof entry !== 'object' || Array.isArray(entry))) throw new Error('Invalid legacy logs');
      return { entries: bounded(value.map((entry, index) => ({ ...entry, seq: index + 1 }))), lastSequence: value.length };
    }
    if (!value || !Array.isArray(value.entries) || !Number.isSafeInteger(value.lastSequence) || value.lastSequence < 0) throw new Error('Invalid log snapshot');
    let previous = value.entries.length ? value.entries[0]?.seq - 1 : value.lastSequence;
    for (const entry of value.entries) {
      if (!entry || !Number.isSafeInteger(entry.seq) || entry.seq < 1 || entry.seq !== previous + 1) throw new Error('Invalid log sequence');
      previous = entry.seq;
    }
    if (previous !== value.lastSequence) throw new Error('Invalid log high-water mark');
    return { entries: bounded(value.entries), lastSequence: value.lastSequence };
  }
  function write(value) {
    while (Buffer.byteLength(JSON.stringify(value)) > MAX_BYTES && value.entries.length) value.entries.shift();
    store.write(LOG_PATH, value);
  }
  function append(entry) {
    const value = read();
    if (value.lastSequence === Number.MAX_SAFE_INTEGER) throw new Error('Log sequence exhausted');
    const seq = value.lastSequence + 1;
    const entries = bounded([...value.entries, { ...entry, receivedAt: new Date().toISOString(), seq }]);
    write({ entries, lastSequence: seq });
    return seq;
  }
  function clear() {
    write({ entries: [], lastSequence: read().lastSequence });
  }
  function after(cursor) {
    const value = read();
    if (cursor > value.lastSequence) {
      const error = new Error('Cursor exceeds stored sequence; possible reset or rollback');
      error.status = 409;
      throw error;
    }
    const oldestCursor = value.entries[0]?.seq || value.lastSequence + 1;
    return { entries: value.entries.filter(entry => entry.seq > cursor), nextCursor: value.lastSequence, oldestCursor, gap: cursor < oldestCursor - 1 };
  }
  return { read, append, clear, after };
}

module.exports = { createLogs };
