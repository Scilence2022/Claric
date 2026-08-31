#!/usr/bin/env node
/**
 * publish-addin — push the built add-in bundle to Scilence2022/claric-addin
 * (the GitHub Pages repo that hosts taskpane.html + manifest.xml at
 * https://scilence2022.github.io/claric-addin/).
 *
 * Flow:
 *  1. Build production bundle into ./dist/ (icons + webpack)
 *  2. Generate manifest.xml from .env (the manifest must point at the
 *     Pages host: scilence2022.github.io, not localhost, before we ship).
 *  3. Sync dist/* + manifest.xml into the local working clone of
 *     Scilence2022/claric-addin and git push to main.
 *
 * Configuration via env:
 *   CLARIC_ADDIN_REPO  Default: Scilence2022/claric-addin
 *   CLARIC_ADDIN_LOCAL Default: /tmp/claric-addin-bootstrap
 *     (a `git clone` of the repo; created once with a README to seed main)
 *
 * Exit codes:
 *   0 success, 1 build/manifest failure, 2 push failure.
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const rootDir = path.resolve(__dirname, '..');
const ADO_REPO = process.env.CLARIC_ADDIN_REPO || 'Scilence2022/claric-addin';
const ADO_LOCAL = process.env.CLARIC_ADDIN_LOCAL || '/tmp/claric-addin-bootstrap';

function run(cmd, args, opts = {}) {
  const r = spawnSync(cmd, args, { stdio: 'inherit', cwd: opts.cwd || rootDir, ...opts });
  if (r.status !== 0) process.exit(r.status || 1);
}

function runCapture(cmd, args, opts = {}) {
  const r = spawnSync(cmd, args, { cwd: opts.cwd || rootDir, encoding: 'utf8', ...opts });
  if (r.status !== 0) {
    process.stderr.write(r.stderr || '');
    process.exit(r.status || 1);
  }
  return r.stdout.trim();
}

function ensureClone() {
  if (fs.existsSync(path.join(ADO_LOCAL, '.git'))) return;
  console.log(`[publish-addin] cloning ${ADO_REPO} into ${ADO_LOCAL}`);
  fs.mkdirSync(path.dirname(ADO_LOCAL), { recursive: true });
  runCapture('git', ['clone', `https://github.com/${ADO_REPO}.git`, ADO_LOCAL]);
  // Seed main if the clone is empty (the Pages repo's main must exist before
  // GH Pages can be enabled, and an empty clone has no branch checked out).
  const heads = runCapture('git', ['-C', ADO_LOCAL, 'branch', '-a']);
  if (!heads.includes('main')) {
    runCapture('git', ['-C', ADO_LOCAL, 'checkout', '-b', 'main']);
    fs.writeFileSync(path.join(ADO_LOCAL, 'README.md'),
      '# claric-addin\n\nStatic hosting for the Claric Word add-in bundle served to Office hosts.\n');
    runCapture('git', ['-C', ADO_LOCAL, 'add', '.']);
    runCapture('git', ['-C', ADO_LOCAL,
      '-c', 'user.email=ci@scilence2022.com',
      '-c', 'user.name=claric-publish',
      'commit', '-m', 'chore: bootstrap repo with README']);
    runCapture('git', ['-C', ADO_LOCAL, 'push', '-u', 'origin', 'main']);
  } else {
    runCapture('git', ['-C', ADO_LOCAL, 'checkout', 'main']);
  }
}

function sync() {
  const distDir = path.join(rootDir, 'dist');
  if (!fs.existsSync(distDir)) {
    console.error('[publish-addin] dist/ missing — run npm run build first');
    process.exit(1);
  }
  const manifestPath = path.join(rootDir, 'manifest.xml');
  if (!fs.existsSync(manifestPath)) {
    console.error('[publish-addin] manifest.xml missing — run npm run manifest first');
    process.exit(1);
  }
  // Clean the Pages repo's tracked files except .git and README, then copy
  // the fresh bundle + manifest. Keeps the bootstrap README so contributors
  // know what the repo is for.
  runCapture('git', ['-C', ADO_LOCAL, 'rm', '-rf', '--cached', '.']);
  // Remove everything except .git and the README from the working tree.
  for (const entry of fs.readdirSync(ADO_LOCAL)) {
    if (entry === '.git' || entry === 'README.md') continue;
    fs.rmSync(path.join(ADO_LOCAL, entry), { recursive: true, force: true });
  }
  // Copy dist/* (HTML, JS, assets/) + manifest.xml.
  fs.cpSync(distDir, ADO_LOCAL, { recursive: true });
  fs.copyFileSync(manifestPath, path.join(ADO_LOCAL, 'manifest.xml'));
}

function push() {
  runCapture('git', ['-C', ADO_LOCAL, 'add', '-A']);
  const status = runCapture('git', ['-C', ADO_LOCAL, 'status', '--porcelain']);
  if (!status) {
    console.log('[publish-addin] no changes — Pages already current');
    return;
  }
  // Use a generic CI identity so commits land even if the developer didn't
  // configure git user.* globally.
  runCapture('git', ['-C', ADO_LOCAL,
    '-c', 'user.email=ci@scilence2022.com',
    '-c', 'user.name=claric-publish',
    'commit', '-m', 'chore: deploy add-in bundle']);
  runCapture('git', ['-C', ADO_LOCAL, 'push', 'origin', 'main']);
  console.log('[publish-addin] pushed to main');
}

console.log('[publish-addin] switching manifest host → store…');
run('npm', ['run', 'manifest:store']);
console.log('[publish-addin] building…');
run('npm', ['run', 'build']);
console.log('[publish-addin] regenerating manifest (post-build icon hash)…');
run('npm', ['run', 'manifest']);
ensureClone();
sync();
push();
console.log('[publish-addin] done — Pages will rebuild in ~30s.');
