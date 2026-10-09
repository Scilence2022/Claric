const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { syncPagesBundle, RETENTION_MS, RETENTION_FILE } = require('../scripts/sync-pages-bundle.cjs');
const { calculateHash } = require('../scripts/verify-build.cjs');

let root;
let source;
let target;
let manifest;
const start = Date.parse('2026-10-09T00:00:00Z');
const DAY = 24 * 60 * 60 * 1000;

function write(dir, name, content) {
    const file = path.join(dir, name);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, content);
}

function release(hash) {
    fs.rmSync(source, { recursive: true, force: true });
    fs.mkdirSync(source);
    write(source, 'taskpane.js', `runtime-${hash}`);
    write(source, 'taskpane.html', `entry-${hash}`);
    write(source, `agent-actions.${hash}.js`, `chunk-${hash}`);
    write(source, `agent-actions.${hash}.js.LICENSE.txt`, `license-${hash}`);
    write(source, 'pdfjs/cmaps/map.bcmap', `map-${hash}`);
    write(source, 'build-info.json', JSON.stringify({ hash }));
}

function sync(now = start) {
    return syncPagesBundle({ distDir: source, targetDir: target, manifestPath: manifest, now });
}

beforeEach(() => {
    root = fs.mkdtempSync(path.join(__dirname, '.pages-sync-'));
    source = path.join(root, 'dist');
    target = path.join(root, 'checkout');
    manifest = path.join(root, 'manifest.xml');
    fs.mkdirSync(target);
    write(root, 'manifest.xml', 'current manifest');
    write(target, '.git/config', 'git metadata');
    write(target, 'README.md', 'hosted add-in');
    release('11111111');
});

afterEach(() => { fs.rmSync(root, { recursive: true, force: true }); });

test('deployment A to B keeps A chunks and licenses while replacing runtime, HTML, and native assets', () => {
    sync();
    release('22222222');
    const before = calculateHash(source);
    expect(sync(start + DAY)).toEqual({ current: 2, retained: 2 });
    expect(fs.readFileSync(path.join(target, 'taskpane.js'), 'utf8')).toBe('runtime-22222222');
    expect(fs.readFileSync(path.join(target, 'taskpane.html'), 'utf8')).toBe('entry-22222222');
    expect(fs.readFileSync(path.join(target, 'pdfjs/cmaps/map.bcmap'), 'utf8')).toBe('map-22222222');
    expect(fs.readFileSync(path.join(target, 'agent-actions.11111111.js'), 'utf8')).toBe('chunk-11111111');
    expect(fs.readFileSync(path.join(target, 'agent-actions.11111111.js.LICENSE.txt'), 'utf8')).toBe('license-11111111');
    expect(fs.readFileSync(path.join(target, '.git/config'), 'utf8')).toBe('git metadata');
    expect(fs.readFileSync(path.join(target, 'README.md'), 'utf8')).toBe('hosted add-in');
    expect(fs.readFileSync(path.join(target, 'manifest.xml'), 'utf8')).toBe('current manifest');
    expect(calculateHash(source)).toBe(before);
    expect(fs.existsSync(path.join(source, RETENTION_FILE))).toBe(false);
});

test('prunes historical chunks only after 30 days while refreshing current chunk retention', () => {
    sync();
    release('22222222');
    sync(start + DAY);
    sync(start + RETENTION_MS);
    expect(fs.existsSync(path.join(target, 'agent-actions.11111111.js'))).toBe(true);
    sync(start + RETENTION_MS + 1);
    expect(fs.existsSync(path.join(target, 'agent-actions.11111111.js'))).toBe(false);
    expect(fs.existsSync(path.join(target, 'agent-actions.11111111.js.LICENSE.txt'))).toBe(false);
    expect(fs.existsSync(path.join(target, 'agent-actions.22222222.js'))).toBe(true);
    const info = JSON.parse(fs.readFileSync(path.join(target, RETENTION_FILE), 'utf8'));
    expect(info.assets['agent-actions.22222222.js']).toBe(start + RETENTION_MS + 1);
    expect(info.assets['agent-actions.11111111.js']).toBeUndefined();
});

test('gives pre-metadata immutable chunks a full retention window, and removes stale mutable files', () => {
    write(target, 'old-tools.abcdef12.js', 'existing chunk');
    write(target, 'old-tools.abcdef12.js.LICENSE.txt', 'existing license');
    write(target, 'stale.js', 'mutable stale file');
    write(target, 'old-assets/icon.png', 'stale asset');
    sync();
    expect(fs.existsSync(path.join(target, 'old-tools.abcdef12.js'))).toBe(true);
    expect(fs.existsSync(path.join(target, 'stale.js'))).toBe(false);
    expect(fs.existsSync(path.join(target, 'old-assets'))).toBe(false);
    sync(start + RETENTION_MS);
    expect(fs.existsSync(path.join(target, 'old-tools.abcdef12.js'))).toBe(true);
    sync(start + RETENTION_MS + 1);
    expect(fs.existsSync(path.join(target, 'old-tools.abcdef12.js'))).toBe(false);
});

test.each(['{invalid', JSON.stringify({ version: 99 }), JSON.stringify({ version: 1, assets: [] })])('recovers conservatively from invalid retention metadata: %s', (data) => {
    write(target, 'old-tools.abcdef12.js', 'existing chunk');
    write(target, RETENTION_FILE, data);
    sync();
    const info = JSON.parse(fs.readFileSync(path.join(target, RETENTION_FILE), 'utf8'));
    expect(info.assets['old-tools.abcdef12.js']).toBe(start);
});

test('does not trust traversal names or invalid timestamps from metadata', () => {
    write(target, 'old-tools.abcdef12.js', 'existing chunk');
    write(target, RETENTION_FILE, JSON.stringify({ version: 1, assets: { '../../outside.js': 1, 'old-tools.abcdef12.js': start + DAY } }));
    sync();
    const info = JSON.parse(fs.readFileSync(path.join(target, RETENTION_FILE), 'utf8'));
    expect(info.assets['../../outside.js']).toBeUndefined();
    expect(info.assets['old-tools.abcdef12.js']).toBe(start);
});

test('rejects immutable filename collisions before changing the publishing checkout', () => {
    sync();
    write(source, 'agent-actions.11111111.js', 'changed without a new filename');
    write(source, 'taskpane.js', 'new runtime');
    expect(() => sync(start + DAY)).toThrow('Immutable chunk changed');
    expect(fs.readFileSync(path.join(target, 'taskpane.js'), 'utf8')).toBe('runtime-11111111');
});

test('rejects source symlinks and reserved publishing paths before removing target assets', () => {
    sync();
    fs.symlinkSync(manifest, path.join(source, 'unsafe.js'));
    expect(() => sync()).toThrow('regular bundle file');
    fs.unlinkSync(path.join(source, 'unsafe.js'));
    write(source, '.git/config', 'must not overwrite git');
    expect(() => sync()).toThrow('reserved publishing files');
    expect(fs.readFileSync(path.join(target, '.git/config'), 'utf8')).toBe('git metadata');
});

test('rejects symlinked historical chunks and retention files', () => {
    fs.symlinkSync(manifest, path.join(target, 'old-tools.abcdef12.js'));
    expect(() => sync()).toThrow('regular bundle file');
    fs.unlinkSync(path.join(target, 'old-tools.abcdef12.js'));
    fs.symlinkSync(manifest, path.join(target, RETENTION_FILE));
    expect(() => sync()).toThrow('regular bundle file');
});

test('rejects overlapping checkouts, target-contained manifests, and policies under 30 days', () => {
    expect(() => syncPagesBundle({ distDir: source, targetDir: root, manifestPath: manifest })).toThrow('must not overlap');
    expect(() => syncPagesBundle({ distDir: source, targetDir: source, manifestPath: manifest })).toThrow('must not overlap');
    expect(() => syncPagesBundle({ distDir: source, targetDir: target, manifestPath: manifest, retentionMs: DAY })).toThrow('at least 30 days');
    expect(() => syncPagesBundle({ distDir: source, targetDir: target, manifestPath: path.join(target, 'README.md') })).toThrow('outside the publishing checkout');
});

test('supports the shared CI publishing command and reports invalid CLI usage', () => {
    const script = path.resolve(__dirname, '../scripts/sync-pages-bundle.cjs');
    const result = spawnSync(process.execPath, [script, source, target, manifest], { encoding: 'utf8' });
    expect(result.status).toBe(0);
    expect(result.stdout).toContain('retained 0 historical chunk asset(s)');
    const invalid = spawnSync(process.execPath, [script], { encoding: 'utf8' });
    expect(invalid.status).toBe(1);
    expect(invalid.stderr).toContain('Usage:');
});
