const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { checkCoverage } = require('../scripts/check-coverage.cjs');

const METRICS = ['statements', 'branches', 'functions', 'lines'];

function counts(covered = 100, total = 100) {
  return Object.fromEntries(METRICS.map((metric) => [metric, { covered, total, pct: 100 }]));
}

function summary() {
  return {
    total: counts(200, 200),
    '/project/src/lib/example.js': counts(),
    '/project/src/taskpane/example.js': counts(),
  };
}

function thresholds(min = 0) {
  return [{ name: 'all', match: () => true, min: Object.fromEntries(METRICS.map((metric) => [metric, min])) }];
}

describe('check-coverage.cjs', () => {
  test('accepts valid coverage and threshold boundaries', () => {
    expect(checkCoverage(summary())).toEqual([]);
    expect(checkCoverage(summary(), thresholds(100))).toEqual([]);
    expect(checkCoverage(summary(), thresholds(0))).toEqual([]);
  });

  test.each([null, [], 1, 'summary', {}, { total: null }, { total: [] }])('rejects invalid summary shape: %j', (value) => {
    expect(checkCoverage(value).length).toBeGreaterThan(0);
  });

  test('requires matching files for every group, not just a reported total', () => {
    const value = summary();
    delete value['/project/src/taskpane/example.js'];
    expect(checkCoverage(value)).toContain('group "src/taskpane/": no files matched');
    expect(checkCoverage({ total: counts() })).toHaveLength(3);
  });

  test('supports Windows file paths', () => {
    const value = summary();
    value['C:\\project\\src\\lib\\example.js'] = value['/project/src/lib/example.js'];
    delete value['/project/src/lib/example.js'];
    expect(checkCoverage(value)).toEqual([]);
  });

  describe.each(METRICS)('%s', (metric) => {
    test.each([undefined, null, [], {}])('rejects missing or malformed metric: %j', (invalid) => {
      const value = summary();
      value['/project/src/lib/example.js'][metric] = invalid;
      expect(checkCoverage(value).join('\n')).toContain(`file "/project/src/lib/example.js" ${metric}`);
    });

    test.each([undefined, null, '100', NaN, Infinity, -Infinity, -1])('rejects invalid counts: %s', (invalid) => {
      for (const field of ['covered', 'total']) {
        const value = summary();
        value['/project/src/lib/example.js'][metric][field] = invalid;
        expect(checkCoverage(value).join('\n')).toContain('requires finite counts');
      }
    });

    test('rejects covered greater than total, including in the reported total', () => {
      for (const file of ['total', '/project/src/lib/example.js']) {
        const value = summary();
        value[file][metric].covered = 201;
        expect(checkCoverage(value).join('\n')).toContain('0 <= covered <= total');
      }
    });

    test('allows a zero-count file but rejects a zero-count group even at a zero threshold', () => {
      const value = summary();
      value['/project/src/lib/example.js'][metric] = { covered: 0, total: 0 };
      expect(checkCoverage(value, thresholds())).toEqual([]);
      value['/project/src/taskpane/example.js'][metric] = { covered: 0, total: 0 };
      expect(checkCoverage(value, thresholds()).join('\n')).toContain('total > 0');
    });

    test('rejects aggregate numeric overflow', () => {
      const value = summary();
      for (const file of Object.keys(value)) value[file][metric] = { total: Number.MAX_VALUE, covered: 0 };
      expect(checkCoverage(value, thresholds()).join('\n')).toContain('total > 0');
    });

    test('recomputes percentages from counts instead of trusting pct or the total entry', () => {
      const value = summary();
      value['/project/src/lib/example.js'][metric].covered = 0;
      expect(checkCoverage(value).join('\n')).toContain(`group "src/lib/" ${metric}: 0.00%`);
    });

    test.each([undefined, null, '70', NaN, Infinity, -Infinity, -1, 101])('rejects invalid threshold: %s', (invalid) => {
      const gates = thresholds();
      gates[0].min[metric] = invalid;
      expect(checkCoverage(summary(), gates).join('\n')).toContain('threshold must be a finite number between 0 and 100');
    });
  });

  test.each([null, [], {}, [null], [{ name: '', match: () => true, min: {} }], [{ name: 'all', min: {} }]])('rejects malformed threshold configuration', (value) => {
    expect(checkCoverage(summary(), value).length).toBeGreaterThan(0);
  });

  test('rejects unknown threshold metrics', () => {
    const gates = thresholds();
    gates[0].min.typo = 50;
    expect(checkCoverage(summary(), gates)).toContain('group "all": unknown threshold metric');
  });

  test('reports all shortfalls deterministically', () => {
    const value = summary();
    for (const file of Object.keys(value)) value[file] = counts(0, 100);
    const failures = checkCoverage(value);
    expect(failures).toHaveLength(12);
    expect(failures[0]).toBe('group "global" statements: 0.00% < required 70%');
    expect(failures).toEqual(checkCoverage(value));
  });

  describe('CLI', () => {
    let root;
    let summaryPath;
    let scriptPath;

    beforeEach(() => {
      root = fs.mkdtempSync(path.join(__dirname, '.coverage-test-'));
      fs.mkdirSync(path.join(root, 'scripts'));
      fs.mkdirSync(path.join(root, 'coverage'));
      scriptPath = path.join(root, 'scripts', 'check-coverage.cjs');
      summaryPath = path.join(root, 'coverage', 'coverage-summary.json');
      fs.copyFileSync(path.join(__dirname, '..', 'scripts', 'check-coverage.cjs'), scriptPath);
    });

    afterEach(() => {
      fs.rmSync(root, { recursive: true, force: true });
    });

    function run() {
      return spawnSync(process.execPath, [scriptPath], { encoding: 'utf8' });
    }

    test('exits zero only for passing coverage', () => {
      fs.writeFileSync(summaryPath, JSON.stringify(summary()));
      const result = run();
      expect(result.status).toBe(0);
      expect(result.stdout).toContain('Coverage gates met.');
      expect(result.stderr).toBe('');
    });

    test.each([undefined, '{broken', 'null', '[]', '{}', '{"total":{}}'])('fails for missing, invalid JSON or invalid shape: %s', (input) => {
      if (input !== undefined) fs.writeFileSync(summaryPath, input);
      const result = run();
      expect(result.status).toBe(1);
      expect(result.stderr).toContain('Coverage gates not met:');
      expect(result.stdout).not.toContain('Coverage gates met.');
    });

    test('rejects JSON numbers that overflow to Infinity', () => {
      fs.writeFileSync(summaryPath, JSON.stringify(summary()).replace('200', '1e400'));
      expect(run().status).toBe(1);
    });

    test('exits nonzero for a genuine threshold shortfall', () => {
      const value = summary();
      value['/project/src/lib/example.js'] = counts(1, 100);
      fs.writeFileSync(summaryPath, JSON.stringify(value));
      const result = run();
      expect(result.status).toBe(1);
      expect(result.stderr).toContain('required');
    });
  });
});
