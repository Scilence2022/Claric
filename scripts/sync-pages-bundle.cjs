#!/usr/bin/env node
/** Keep immutable chunks addressable for Word panes opened before a deploy. */
const fs = require('fs');
const path = require('path');

const RETENTION_MS = 30 * 24 * 60 * 60 * 1000;
const RETENTION_FILE = 'chunk-retention.json';
const HASHED_CHUNK = /^[A-Za-z0-9][A-Za-z0-9_.-]*\.[a-f0-9]{8,64}\.(?:js|mjs)(?:\.LICENSE\.txt)?$/;

function assertRegular(file) {
    if (!fs.lstatSync(file).isFile()) throw new Error(`Expected a regular bundle file: ${file}`);
}

function listRegularFiles(directory, prefix = '') {
    const files = [];
    if (!fs.lstatSync(directory).isDirectory()) throw new Error(`Expected a bundle directory: ${directory}`);
    for (const name of fs.readdirSync(directory)) {
        const file = path.join(directory, name);
        const relative = prefix ? `${prefix}/${name}` : name;
        const stat = fs.lstatSync(file);
        if (stat.isDirectory()) files.push(...listRegularFiles(file, relative));
        else { assertRegular(file); files.push(relative); }
    }
    return files;
}

function readRetention(targetDir) {
    const file = path.join(targetDir, RETENTION_FILE);
    if (!fs.existsSync(file)) return {};
    assertRegular(file);
    try {
        const info = JSON.parse(fs.readFileSync(file, 'utf8'));
        return info?.version === 1 && info.assets && typeof info.assets === 'object'
            && !Array.isArray(info.assets) ? info.assets : {};
    } catch (_error) {
        // Conservatively start a new retention window for pre-metadata files.
        return {};
    }
}

/**
 * Source dist stays pristine: historical chunks and retention metadata are
 * added only to the publishing checkout, after production build verification.
 */
function syncPagesBundle({ distDir, targetDir, manifestPath, now = Date.now(), retentionMs = RETENTION_MS }) {
    const source = path.resolve(distDir);
    const target = path.resolve(targetDir);
    const overlaps = (a, b) => {
        const relative = path.relative(a, b);
        return !relative || (relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative));
    };
    if (overlaps(source, target) || overlaps(target, source)) throw new Error('Bundle source and publishing checkout must not overlap.');
    if (!Number.isFinite(now) || now < 0 || !Number.isFinite(retentionMs) || retentionMs < RETENTION_MS) {
        throw new Error('Historical chunks must be retained for at least 30 days.');
    }
    const files = listRegularFiles(source);
    if (files.some((file) => file === RETENTION_FILE || file === 'README.md' || file === '.git' || file.startsWith('.git/'))) {
        throw new Error('Bundle contains reserved publishing files.');
    }
    if (overlaps(target, path.resolve(manifestPath))) throw new Error('Source manifest must be outside the publishing checkout.');
    assertRegular(manifestPath);
    if (!fs.lstatSync(target).isDirectory()) throw new Error('Publishing checkout must be a regular directory.');
    const previous = readRetention(target);
    const currentChunks = new Set(files.filter((file) => HASHED_CHUNK.test(file)));
    const assets = Object.create(null);
    const retained = new Set();
    for (const name of fs.readdirSync(target)) {
        if (!HASHED_CHUNK.test(name)) continue;
        const file = path.join(target, name);
        assertRegular(file);
        if (currentChunks.has(name) && !fs.readFileSync(file).equals(fs.readFileSync(path.join(source, name)))) {
            throw new Error(`Immutable chunk changed without a new filename: ${name}`);
        }
        const recorded = previous[name];
        const lastPublished = Number.isFinite(recorded) && recorded >= 0 && recorded <= now ? recorded : now;
        if (currentChunks.has(name) || now - lastPublished <= retentionMs) {
            retained.add(name);
            assets[name] = currentChunks.has(name) ? now : lastPublished;
        }
    }
    for (const name of currentChunks) assets[name] = now;
    for (const name of fs.readdirSync(target)) {
        if (name === '.git' || name === 'README.md' || retained.has(name)) continue;
        fs.rmSync(path.join(target, name), { recursive: true, force: true });
    }
    fs.cpSync(source, target, { recursive: true });
    fs.copyFileSync(manifestPath, path.join(target, 'manifest.xml'));
    fs.writeFileSync(path.join(target, RETENTION_FILE), `${JSON.stringify({ version: 1, assets }, null, 2)}\n`);
    return { current: currentChunks.size, retained: [...retained].filter((name) => !currentChunks.has(name)).length };
}

if (require.main === module) {
    try {
        const [distDir, targetDir, manifestPath] = process.argv.slice(2);
        if (!distDir || !targetDir || !manifestPath) throw new Error('Usage: sync-pages-bundle.cjs <dist> <checkout> <manifest>');
        const result = syncPagesBundle({ distDir, targetDir, manifestPath });
        console.log(`[deploy] synced bundle; retained ${result.retained} historical chunk asset(s) for open Word panes`);
    } catch (error) {
        console.error(`[deploy] ${error.message}`);
        process.exitCode = 1;
    }
}

module.exports = { syncPagesBundle, RETENTION_MS, RETENTION_FILE };
