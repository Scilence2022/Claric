#!/usr/bin/env node
/**
 * manifest-mode — switch .env between "local dev" and "store" host and
 * regenerate manifest.xml in one command. Avoids the silent footgun of
 * editing .env by hand and forgetting to re-run manifest.
 *
 *   node scripts/manifest-mode.cjs local   # HOST=localhost, PORT=3001
 *   node scripts/manifest-mode.cjs store   # HOST=scilence2022.github.io, PORT=443
 *
 * Reads the current .env, swaps HOST / HOST_PORT / PROTOCOL only, and
 * writes the other env lines back untouched. Idempotent.
 */
const fs = require('fs');
const path = require('path');

const envPath = path.resolve(__dirname, '..', '.env');
const mode = process.argv[2];
if (mode !== 'local' && mode !== 'store') {
  console.error('usage: node scripts/manifest-mode.cjs <local|store>');
  process.exit(64);
}

const PROFILES = {
  local: { HOST: 'localhost', HOST_PORT: '3001', PROTOCOL: 'https' },
  store: { HOST: 'scilence2022.github.io', HOST_PORT: '443', PROTOCOL: 'https' },
};
const target = PROFILES[mode];

let envBody;
try {
  envBody = fs.readFileSync(envPath, 'utf8');
} catch (e) {
  console.error(`.env not found at ${envPath} — copy .env.example first.`);
  process.exit(1);
}

const lines = envBody.split(/\r?\n/);
const keys = new Set(Object.keys(target));
let changed = 0;
for (let i = 0; i < lines.length; i++) {
  const m = lines[i].match(/^([A-Z0-9_]+)=(.*)$/);
  if (!m || !keys.has(m[1])) continue;
  if (target[m[1]] === m[2]) continue;
  lines[i] = `${m[1]}=${target[m[1]]}`;
  changed++;
}

fs.writeFileSync(envPath, lines.join('\n'));
console.log(`[manifest-mode] ${mode}: ${changed} field(s) updated`);
const { spawnSync } = require('child_process');
const r = spawnSync('npm', ['run', 'manifest'], { stdio: 'inherit' });
process.exit(r.status || 0);
