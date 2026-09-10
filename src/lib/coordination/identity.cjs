const crypto = require('crypto');

const ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/;
function normalizeId(value, field = 'id') {
  if (typeof value !== 'string' || !ID_PATTERN.test(value)) throw new Error(`Invalid ${field}`);
  return value;
}
function createClientId() { return crypto.randomUUID(); }
module.exports = { normalizeId, createClientId };
