#!/usr/bin/env node
/**
 * sideload-addin — install manifest.xml into Word's local manifest
 * folder so the add-in shows up in Home > Add-ins on next launch.
 *
 * macOS is the only platform with a programmatic sideload path: Word reads
 * manifests from `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
 * on startup. On Windows and Word on the web, sideloading is UI-only — the
 * script prints the menu paths and exits.
 *
 *   node scripts/sideload-addin.cjs           # copy manifest.xml into wef/
 *   node scripts/sideload-addin.cjs --remove  # clear wef/ of any prior Claric manifest
 *
 * Idempotent: re-running replaces the previous manifest.xml. The script
 * refuses to touch wef/ if it doesn't exist (no create-from-scratch —
 * Word installs that path on first launch with a sideloaded add-in).
 */
const fs = require('fs');
const path = require('path');
const os = require('os');

const rootDir = path.resolve(__dirname, '..');
const manifestSrc = path.join(rootDir, 'manifest.xml');

if (!fs.existsSync(manifestSrc)) {
  console.error(`manifest.xml missing at ${manifestSrc} — run \`npm run manifest\` first.`);
  process.exit(1);
}

const remove = process.argv.includes('--remove');

function printUiPaths(platform) {
  if (platform === 'win32') {
    console.log(`
Windows has no CLI sideload path. In Word:

  File → Options → Trust Center → Trust Center Settings →
  Trusted Add-in Catalogs → Add the folder containing manifest.xml →
  check "Show in Menu" → OK. Restart Word.

Then: Home → Add-ins → Manage My Add-ins → "Shared Folder" tab →
pick manifest.xml.
`);
  } else {
    console.log(`
Word on the web has no CLI sideload path. In Word on the web:

  Insert → Add-ins → Manage My Add-ins → "Upload My Add-in" →
  pick manifest.xml.
`);
  }
}

if (process.platform === 'darwin') {
  const wef = path.join(os.homedir(),
    'Library', 'Containers', 'com.microsoft.Word', 'Data', 'Documents', 'wef');
  if (remove) {
    if (!fs.existsSync(wef)) {
      console.log(`wef/ missing at ${wef} — nothing to remove.`);
      process.exit(0);
    }
    for (const entry of fs.readdirSync(wef)) {
      fs.rmSync(path.join(wef, entry), { recursive: true, force: true });
    }
    console.log(`[sideload-addin] cleared ${wef}`);
    console.log('Restart Word for the change to take effect.');
    process.exit(0);
  }
  if (!fs.existsSync(wef)) {
    console.error(`wef/ missing at ${wef}.`);
    console.error('Sideload a manifest once via the UI (Insert → Add-ins → My Add-ins → Developer),');
    console.error('then re-run this script. Word creates wef/ on first sideload.');
    process.exit(1);
  }
  // Remove any prior claric add-in manifest (so re-running isn't a stale
  // ghost) by Id-matching the existing files.
  const idMatch = fs.readFileSync(manifestSrc, 'utf8').match(/<Id>([^<]+)<\/Id>/);
  const newId = idMatch ? idMatch[1].trim() : null;
  for (const entry of fs.readdirSync(wef)) {
    const fp = path.join(wef, entry);
    if (entry === '.DS_Store' || !fs.statSync(fp).isFile()) continue;
    if (!entry.endsWith('.xml')) continue;
    try {
      const body = fs.readFileSync(fp, 'utf8');
      const m = body.match(/<Id>([^<]+)<\/Id>/);
      if (m && newId && m[1].trim() === newId) {
        fs.rmSync(fp);
        console.log(`[sideload-addin] removed prior manifest ${entry}`);
      }
    } catch (_e) { /* skip unreadable */ }
  }
  const dest = path.join(wef, 'manifest.xml');
  fs.copyFileSync(manifestSrc, dest);
  console.log(`[sideload-addin] copied ${manifestSrc}`);
  console.log(`               → ${dest}`);
  console.log('Restart Word, then Home → Add-ins → Claric.');
} else {
  printUiPaths(process.platform);
}
