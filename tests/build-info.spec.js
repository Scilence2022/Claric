/**
 * Tests for the dist/build-info.json emitter produced by webpack.config.cjs
 * (the inline plugin keyed 'claric-build-info' that runs in production mode
 * after webpack emits assets). The plugin only writes the file under
 * production mode, so this spec needs a pre-existing dist/build-info.json.
 *
 * The plugin cannot run inside Jest, so we assert structural facts about
 * whatever build-info.json the developer has on disk (left over from a
 * local `npm run build` or CI artifact). The test is skipped gracefully if
 * the file is missing — this prevents the spec from failing on machines
 * where dist/ is gitignored and hasn't been built yet.
 */

const fs = require('fs');
const path = require('path');

const BUILD_INFO_PATH = path.resolve(__dirname, '..', 'dist', 'build-info.json');
const PACKAGE_JSON_PATH = path.resolve(__dirname, '..', 'package.json');

const HASH_REGEX = /^[0-9a-f]{12}$/;
const ISO_8601_REGEX = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$/;

const exists = fs.existsSync(BUILD_INFO_PATH);
const describeIfBuilt = exists ? describe : describe.skip;

describe('dist/build-info.json', () => {
    if (!exists) {
        it('is not present — run `npm run build` first', () => {
            console.warn(
                `[build-info] ${BUILD_INFO_PATH} missing; test skipped. ` +
                'Run `npm run build` to produce the artifact.'
            );
        });
        return;
    }

    let info;
    let pkg;

    beforeAll(() => {
        info = JSON.parse(fs.readFileSync(BUILD_INFO_PATH, 'utf8'));
        pkg = JSON.parse(fs.readFileSync(PACKAGE_JSON_PATH, 'utf8'));
    });

    describeIfBuilt('shape and field format', () => {
        it('exposes exactly the documented keys', () => {
            expect(Object.keys(info).sort()).toEqual(
                ['appVersion', 'builtAt', 'hash', 'mode'].sort()
            );
        });

        it('hash is a 12-character lowercase hex string', () => {
            expect(typeof info.hash).toBe('string');
            expect(info.hash).toMatch(HASH_REGEX);
        });

        it('builtAt is an ISO-8601 UTC timestamp', () => {
            expect(typeof info.builtAt).toBe('string');
            expect(info.builtAt).toMatch(ISO_8601_REGEX);
            // Sanity: parseable as a real Date.
            expect(Number.isNaN(Date.parse(info.builtAt))).toBe(false);
        });

        it('mode is the literal string "production"', () => {
            expect(info.mode).toBe('production');
        });

        it('appVersion mirrors package.json#version', () => {
            expect(info.appVersion).toBe(String(pkg.version));
        });
    });
});
