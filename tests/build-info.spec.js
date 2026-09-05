const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawnSync } = require('child_process');
const {
    REQUIRED_ASSETS, normalizeAssetPath, calculateHash, verifyBuild,
} = require('../scripts/verify-build.cjs');

jest.mock('dotenv', () => ({ config: jest.fn() }));

let rootDir;
let distDir;
let assets;
let info;

function write(name, content) {
    const target = path.join(rootDir, name);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, content);
}

function fixtureHash() {
    const records = Object.entries(assets).map(([name, content]) =>
        `${name}:${crypto.createHash('sha256').update(content).digest('hex')}`
    ).sort();
    return crypto.createHash('sha256').update(records.join('\n')).digest('hex').slice(0, 12);
}

function saveInfo() {
    write('dist/build-info.json', JSON.stringify(info));
}

function cli() {
    return spawnSync(process.execPath, [path.join(rootDir, 'scripts/verify-build.cjs')], {
        cwd: rootDir, encoding: 'utf8',
    });
}

beforeEach(() => {
    rootDir = fs.mkdtempSync(path.resolve(__dirname, '.build-fixture-'));
    distDir = path.join(rootDir, 'dist');
    assets = Object.fromEntries(REQUIRED_ASSETS.map((name) => [name, `fixture: ${name}`]));
    assets['nested/build-info.json'] = '{"nested":true}';
    for (const [name, content] of Object.entries(assets)) write(`dist/${name}`, content);
    write('package.json', JSON.stringify({ version: '1.2.3' }));
    write('scripts/verify-build.cjs', fs.readFileSync(path.resolve(__dirname, '../scripts/verify-build.cjs')));
    info = { appVersion: '1.2.3', builtAt: '2026-09-05T06:00:00.123Z', hash: fixtureHash(), mode: 'production' };
    saveInfo();
});

afterEach(() => {
    jest.restoreAllMocks();
    fs.rmSync(rootDir, { recursive: true, force: true });
});

describe('build verifier without a pre-existing dist', () => {
    test('verifies an independently hashed fixture and exits zero', () => {
        expect(verifyBuild({ rootDir })).toEqual(info);
        expect(calculateHash(distDir)).toBe(fixtureHash());
        const result = cli();
        expect(result.status).toBe(0);
        expect(result.stdout).toContain(`verified production build 1.2.3 (${info.hash})`);
    });

    test.each([
        ['null', null], ['array', []], ['missing field', { hash: 'a'.repeat(12) }],
        ['unknown field', { extra: true }], ['numeric hash', { hash: 123 }],
        ['uppercase hash', { hash: 'A'.repeat(12) }], ['short hash', { hash: 'abc' }],
        ['development mode', { mode: 'development' }], ['numeric version', { appVersion: 123 }],
        ['wrong version', { appVersion: '2.0.0' }], ['invalid date', { builtAt: 'not-a-date' }],
        ['calendar overflow', { builtAt: '2026-02-30T06:00:00Z' }],
        ['invalid month', { builtAt: '2026-13-01T06:00:00Z' }],
        ['invalid hour', { builtAt: '2026-09-05T24:00:00Z' }],
        ['non UTC', { builtAt: '2026-09-05T06:00:00+01:00' }],
        ['numeric date', { builtAt: 0 }],
    ])('rejects %s with exit 1', (_label, value) => {
        info = value && !Array.isArray(value) && _label !== 'missing field' ? { ...info, ...value } : value;
        saveInfo();
        expect(() => verifyBuild({ rootDir })).toThrow();
        expect(cli().status).toBe(1);
    });

    test.each(['dist', 'dist/build-info.json', 'package.json'])('rejects missing %s', (name) => {
        fs.rmSync(path.join(rootDir, name), { recursive: true, force: true });
        expect(cli().status).toBe(1);
    });

    test.each(['dist/build-info.json', 'package.json'])('rejects malformed %s', (name) => {
        write(name, '{');
        expect(cli().status).toBe(1);
    });

    test.each([{}, { version: 123 }, null])('rejects invalid package metadata %j', (pkg) => {
        write('package.json', JSON.stringify(pkg));
        expect(cli().status).toBe(1);
    });

    test.each(REQUIRED_ASSETS)('rejects missing required %s even with self-consistent hash', (name) => {
        fs.unlinkSync(path.join(distDir, name));
        delete assets[name];
        info.hash = fixtureHash();
        saveInfo();
        expect(() => verifyBuild({ rootDir })).toThrow();
        expect(cli().status).toBe(1);
    });

    test.each(['taskpane.js', 'nested/build-info.json'])('detects modified %s', (name) => {
        write(`dist/${name}`, 'changed');
        expect(() => verifyBuild({ rootDir })).toThrow(/hash mismatch/);
        expect(cli().status).toBe(1);
    });

    test('detects missing optional assets and extra assets', () => {
        fs.unlinkSync(path.join(distDir, 'nested/build-info.json'));
        expect(() => verifyBuild({ rootDir })).toThrow(/hash mismatch/);
        write('dist/nested/build-info.json', assets['nested/build-info.json']);
        write('dist/unexpected.js', 'extra');
        expect(() => verifyBuild({ rootDir })).toThrow(/hash mismatch/);
    });

    test('hash ignores only root metadata and is independent of creation order', () => {
        const expected = info.hash;
        write('dist/build-info.json', 'arbitrary metadata');
        expect(calculateHash(distDir)).toBe(expected);
        fs.rmSync(distDir, { recursive: true });
        for (const [name, content] of Object.entries(assets).reverse()) write(`dist/${name}`, content);
        expect(calculateHash(distDir)).toBe(expected);
    });

    test('normalizes Windows and POSIX asset paths identically', () => {
        expect(normalizeAssetPath('assets\\nested\\icon.png')).toBe('assets/nested/icon.png');
        expect(normalizeAssetPath('assets/nested/icon.png')).toBe('assets/nested/icon.png');
    });

    test.each(['file', 'directory', 'broken', 'metadata'])('rejects %s symlinks', (kind) => {
        let link = path.join(distDir, 'linked');
        let target = path.join(distDir, kind === 'directory' ? 'assets' : 'taskpane.js');
        if (kind === 'broken') target = path.join(distDir, 'missing');
        if (kind === 'metadata') {
            link = path.join(distDir, 'build-info.json');
            write('metadata.json', JSON.stringify(info));
            target = path.join(rootDir, 'metadata.json');
            fs.unlinkSync(link);
        }
        fs.symlinkSync(target, link, kind === 'directory' ? 'dir' : 'file');
        expect(() => verifyBuild({ rootDir })).toThrow(/Symbolic links/);
        expect(cli().status).toBe(1);
    });

    test('rejects unreadable assets', () => {
        const file = path.join(distDir, 'taskpane.js');
        fs.chmodSync(file, 0);
        try {
            expect(() => verifyBuild({ rootDir })).toThrow(/Unreadable/);
            expect(cli().status).toBe(1);
        } finally {
            fs.chmodSync(file, 0o644);
        }
    });

    test('propagates asset read failures instead of omitting the file', () => {
        const read = fs.readFileSync;
        jest.spyOn(fs, 'readFileSync').mockImplementation((file, ...args) => {
            if (String(file).endsWith('taskpane.js')) throw new Error('EACCES: permission denied');
            return read(file, ...args);
        });
        expect(() => verifyBuild({ rootDir })).toThrow(/EACCES/);
    });

    test('importing the CLI has no verification or process side effects', () => {
        fs.rmSync(distDir, { recursive: true });
        const result = spawnSync(process.execPath, ['-e', 'require("./scripts/verify-build.cjs")'], {
            cwd: rootDir, encoding: 'utf8',
        });
        expect(result.status).toBe(0);
        expect(result.stdout).toBe('');
        expect(result.stderr).toBe('');
    });
});

describe('webpack build-info emitter', () => {
    function emit({ errors = false, mode = 'production', windows = false } = {}) {
        const config = require('../webpack.config.cjs')({}, { mode });
        const callbacks = {};
        const compiler = { hooks: { done: { tap: (name, callback) => { callbacks[name] = callback; } } } };
        config.plugins[config.plugins.length - 1].apply(compiler);
        callbacks['claric-build-info']({
            hasErrors: () => errors,
            compilation: {
                options: { mode }, outputOptions: { path: distDir },
                getAssets: () => Object.keys(assets).map((name) => ({ name: windows ? name.replace(/\//g, '\\') : name })),
            },
        });
        return config;
    }

    test.each([false, true])('hashes current assets with normalized paths (Windows: %s)', (windows) => {
        const config = emit({ windows });
        const emitted = JSON.parse(fs.readFileSync(path.join(distDir, 'build-info.json'), 'utf8'));
        expect(emitted.hash).toBe(fixtureHash());
        expect(config.output.clean).toBe(true);
        write('package.json', JSON.stringify({ version: emitted.appVersion }));
        expect(verifyBuild({ rootDir })).toEqual(emitted);
    });

    test('does not include stale disk files in compilation hash; verifier rejects them', () => {
        write('dist/stale.js', 'stale');
        emit();
        info = JSON.parse(fs.readFileSync(path.join(distDir, 'build-info.json'), 'utf8'));
        expect(info.hash).toBe(fixtureHash());
        write('package.json', JSON.stringify({ version: info.appVersion }));
        expect(() => verifyBuild({ rootDir })).toThrow(/hash mismatch/);
    });

    test('removes stale success metadata after compilation errors', () => {
        emit({ errors: true });
        expect(fs.existsSync(path.join(distDir, 'build-info.json'))).toBe(false);
        expect(cli().status).toBe(1);
    });

    test('does not emit metadata for development compilation', () => {
        fs.unlinkSync(path.join(distDir, 'build-info.json'));
        emit({ mode: 'development' });
        expect(fs.existsSync(path.join(distDir, 'build-info.json'))).toBe(false);
    });
});
