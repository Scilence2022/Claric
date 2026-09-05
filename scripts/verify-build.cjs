#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const REQUIRED_ASSETS = Object.freeze([
  'taskpane.html', 'taskpane.js', 'commands.html', 'commands.js',
  'pdf.worker.min.mjs', 'assets/icon.svg',
  ...[16, 32, 64, 80, 128].map((size) => `assets/icon-${size}.png`),
]);
const REQUIRED_KEYS = ['appVersion', 'builtAt', 'hash', 'mode'];
const HASH_REGEX = /^[0-9a-f]{12}$/;
const ISO_8601_REGEX = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$/;

function normalizeAssetPath(name) {
  return name.replace(/\\/g, '/');
}

function assertReadable(filePath, directory = false) {
  const stat = fs.lstatSync(filePath);
  if (stat.isSymbolicLink()) throw new Error(`Symbolic links are not allowed: ${filePath}`);
  if (directory ? !stat.isDirectory() : !stat.isFile()) {
    throw new Error(`Expected ${directory ? 'directory' : 'regular file'}: ${filePath}`);
  }
  if (!(stat.mode & 0o444)) throw new Error(`Unreadable asset: ${filePath}`);
  fs.accessSync(filePath, directory ? fs.constants.R_OK | fs.constants.X_OK : fs.constants.R_OK);
}

function readJson(filePath) {
  assertReadable(filePath);
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function collectFiles(dir, baseDir = dir) {
  assertReadable(dir, true);
  const files = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) files.push(...collectFiles(fullPath, baseDir));
    else {
      assertReadable(fullPath);
      files.push({ fullPath, name: normalizeAssetPath(path.relative(baseDir, fullPath)) });
    }
  }
  return files;
}

function calculateHash(distDir) {
  const files = collectFiles(distDir);
  if (new Set(files.map(({ name }) => name)).size !== files.length) throw new Error('Duplicate normalized asset paths');
  const fileHashes = files.filter(({ name }) => name !== 'build-info.json').map(({ name, fullPath }) => {
    const digest = crypto.createHash('sha256')
      .update(fs.readFileSync(fullPath))
      .digest('hex');
    return `${name}:${digest}`;
  }).sort();
  return crypto.createHash('sha256').update(fileHashes.join('\n')).digest('hex').slice(0, 12);
}

function verifyBuild({ rootDir = path.resolve(__dirname, '..') } = {}) {
  const distDir = path.join(rootDir, 'dist');
  assertReadable(distDir, true);
  const info = readJson(path.join(distDir, 'build-info.json'));
  const pkg = readJson(path.join(rootDir, 'package.json'));
  if (!info || typeof info !== 'object' || Array.isArray(info)) {
    throw new Error('build-info.json must contain an object');
  }
  if (Object.keys(info).sort().join(',') !== REQUIRED_KEYS.join(',')) {
    throw new Error(`build-info.json must contain exactly: ${REQUIRED_KEYS.join(', ')}`);
  }
  if (typeof info.hash !== 'string' || !HASH_REGEX.test(info.hash)) {
    throw new Error('build-info.hash must be a 12-character lowercase hexadecimal string');
  }
  if (typeof info.builtAt !== 'string' || !ISO_8601_REGEX.test(info.builtAt)
      || !Number.isFinite(Date.parse(info.builtAt))
      || new Date(info.builtAt).toISOString().slice(0, 19) !== info.builtAt.slice(0, 19)) {
    throw new Error('build-info.builtAt must be a valid ISO-8601 UTC timestamp');
  }
  if (info.mode !== 'production') throw new Error('build-info.mode must be "production"');
  if (!pkg || typeof pkg.version !== 'string' || !pkg.version
      || typeof info.appVersion !== 'string' || info.appVersion !== pkg.version) {
    throw new Error('build-info.appVersion must match package.json.version');
  }
  for (const name of REQUIRED_ASSETS) assertReadable(path.join(distDir, name));
  const hash = calculateHash(distDir);
  if (info.hash !== hash) {
    throw new Error(`build hash mismatch: build-info.json has ${info.hash}, calculated ${hash}`);
  }
  return info;
}

function main(options) {
  try {
    const info = verifyBuild(options);
    console.log(`[verify-build] verified production build ${info.appVersion} (${info.hash})`);
    return 0;
  } catch (error) {
    console.error(`[verify-build] ${error.message}`);
    return 1;
  }
}

if (require.main === module) process.exitCode = main();

module.exports = { REQUIRED_ASSETS, normalizeAssetPath, calculateHash, verifyBuild, main };
