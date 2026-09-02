#!/usr/bin/env node
/**
 * sideload-addin ? install manifest.xml into Word's per-platform sideload
 * location so the add-in loads on next launch / document open.
 *
 *   node scripts/sideload-addin.cjs           # install
 *   node scripts/sideload-addin.cjs --remove  # uninstall
 *
 * - macOS: copies manifest.xml into Word's wef/ container folder (Word
 *   reads it on startup; the folder must already exist ? Word creates it
 *   on the first UI sideload).
 * - Windows: runs installer/windows/Install-Claric.ps1 (or Uninstall on
 *   --remove), which registers the manifest under
 *   HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer and builds a launch
 *   document. Works on consumer Microsoft 365 builds where "Upload My
 *   Add-in" is gone and the trusted-catalog route needs a UNC share.
 *
 * Idempotent: re-running replaces the previous registration.
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawnSync } = require('child_process');

const rootDir = path.resolve(__dirname, '..');
const manifestSrc = path.join(rootDir, 'manifest.xml');

if (!fs.existsSync(manifestSrc)) {
  console.error(`manifest.xml missing at ${manifestSrc} ? run \`npm run manifest\` first.`);
  process.exit(1);
}

const remove = process.argv.includes('--remove');

function printUiPaths(platform) {
  if (platform === 'win32') {
    console.log(`
The Windows installer failed to launch. Manual fallback:

  powershell -ExecutionPolicy Bypass -File installer\\windows\\${remove ? 'Un' : ''}Install-Claric.ps1 -ManifestPath manifest.xml

See installer/windows/README.md for details and manual UI paths.
`);
  } else {
    console.log(`
Word on the web has no CLI sideload path. In Word on the web:

  Insert ? Add-ins ? Manage My Add-ins ? "Upload My Add-in" ?
  pick manifest.xml.
`);
  }
}

function runWindowsInstaller() {
  const script = path.join(rootDir, 'installer', 'windows', `${remove ? 'Un' : ''}Install-Claric.ps1`);
  if (!fs.existsSync(script)) {
    console.error(`installer script missing: ${script}`);
    process.exit(1);
  }
  const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script];
  if (!remove) args.push('-ManifestPath', manifestSrc, '-NoLaunch');
  else args.push('-ManifestPath', manifestSrc);
  const result = spawnSync('powershell.exe', args, { stdio: 'inherit', windowsHide: true });
  if (result.error || result.status !== 0) {
    printUiPaths('win32');
    process.exit(result.status === null ? 1 : result.status);
  }
  console.log('[sideload-addin] Windows install complete.');
  if (!remove) {
    console.log('Open the launch document printed above (or Insert ? Get Add-ins ? My Add-ins ? Claric).');
  }
}

if (process.platform === 'win32') {
  runWindowsInstaller();
} else if (process.platform === 'darwin') {
  const wef = path.join(os.homedir(),
    'Library', 'Containers', 'com.microsoft.Word', 'Data', 'Documents', 'wef');
  if (remove) {
    if (!fs.existsSync(wef)) {
      console.log(`wef/ missing at ${wef} ? nothing to remove.`);
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
    console.error('Sideload a manifest once via the UI (Insert ? Add-ins ? My Add-ins ? Developer),');
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
  console.log(`               ? ${dest}`);
  console.log('Restart Word, then Home ? Add-ins ? Claric.');
} else {
  printUiPaths(process.platform);
}
