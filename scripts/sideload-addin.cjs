#!/usr/bin/env node
/**
 * sideload-addin ? install manifest.xml into Word's per-platform sideload
 * location so the add-in loads on next launch / document open.
 *
 *   node scripts/sideload-addin.cjs           # install
 *   node scripts/sideload-addin.cjs --remove  # uninstall
 *
 * - macOS: runs installer/macos/Install-Claric.sh (or Uninstall on
 *   --remove), which copies the manifest into Word's wef/ container folder
 *   (Word reads it on startup; the script creates the folder if needed) and
 *   builds a launch document. No admin rights, no Node.js required on the
 *   target Mac beyond this script itself.
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
  } else if (platform === 'darwin') {
    console.log(`
The macOS installer failed to launch. Manual fallback:

  bash installer/macos/${remove ? 'Un' : ''}Install-Claric.sh --manifest manifest.xml

See installer/macos/README.md for details and manual UI paths.
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
  console.log(`[sideload-addin] Windows ${remove ? 'removal' : 'install'} complete.`);
  if (!remove) {
    console.log('Open the launch document printed above (or Insert ? Get Add-ins ? My Add-ins ? Claric).');
  }
}

function runMacInstaller() {
  const script = path.join(rootDir, 'installer', 'macos', `${remove ? 'Un' : ''}Install-Claric.sh`);
  if (!fs.existsSync(script)) {
    console.error(`installer script missing: ${script}`);
    process.exit(1);
  }
  const args = [script, '--manifest', manifestSrc];
  if (!remove) args.push('--no-launch');
  const result = spawnSync('bash', args, { stdio: 'inherit' });
  if (result.error || result.status !== 0) {
    printUiPaths('darwin');
    process.exit(result.status === null ? 1 : result.status);
  }
  console.log(`[sideload-addin] macOS ${remove ? 'removal' : 'install'} complete.`);
  if (!remove) {
    console.log('Open the launch document printed above (or Home ? Add-ins ? Claric).');
  }
}

if (process.platform === 'win32') {
  runWindowsInstaller();
} else if (process.platform === 'darwin') {
  runMacInstaller();
} else {
  printUiPaths(process.platform);
}
