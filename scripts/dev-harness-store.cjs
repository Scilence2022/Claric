const fs = require('fs');
const path = require('path');
const { randomUUID } = require('crypto');

const MAX_ENTRIES = 1000;
const MAX_BYTES = 4 * 1024 * 1024;

function bounded(entries) {
  const result = entries.slice(-MAX_ENTRIES);
  let bytes = Buffer.byteLength(JSON.stringify(result));
  while (bytes > MAX_BYTES && result.length) {
    bytes -= Buffer.byteLength(JSON.stringify(result.shift())) + 1;
  }
  return result;
}

function createStore(rootDir) {
  const root = fs.realpathSync(rootDir);
  function resolve(relative) {
    const target = path.resolve(root, relative);
    if (!target.startsWith(`${root}${path.sep}`)) throw new Error('Invalid storage path');
    let current = root;
    for (const segment of path.relative(root, target).split(path.sep)) {
      current = path.join(current, segment);
      try {
        if (fs.lstatSync(current).isSymbolicLink()) throw new Error('Symlink storage is not supported');
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
    }
    return target;
  }
  function read(relative, fallback = []) {
    const target = resolve(relative);
    try {
      if (fs.statSync(target).size > MAX_BYTES) throw new Error('Stored snapshot exceeds size limit');
      return JSON.parse(fs.readFileSync(target, 'utf8'));
    } catch (error) {
      if (error.code === 'ENOENT') return fallback;
      throw error;
    }
  }
  function readArray(relative) {
    const value = read(relative);
    if (!Array.isArray(value)) throw new Error('Stored snapshot must be an array');
    return value;
  }
  function write(relative, value) {
    const data = JSON.stringify(value);
    if (Buffer.byteLength(data) > MAX_BYTES) {
      const error = new Error('Snapshot capacity exceeded');
      error.status = 413;
      throw error;
    }
    const target = resolve(relative);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    const temporary = `${target}.${randomUUID()}.tmp`;
    try {
      fs.writeFileSync(temporary, data, { encoding: 'utf8', flag: 'wx', mode: 0o600 });
      fs.renameSync(temporary, target);
    } catch (error) {
      try { fs.unlinkSync(temporary); } catch (cleanupError) {
        if (cleanupError.code !== 'ENOENT') error.cleanupFailed = true;
      }
      throw error;
    }
  }
  return { read, readArray, write };
}

module.exports = { createStore, bounded, MAX_ENTRIES, MAX_BYTES };
