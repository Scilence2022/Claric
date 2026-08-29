/**
 * Enforces the coverage gates from coverage/coverage-summary.json.
 *
 * Run after `npm run coverage` (jest --coverage with the json-summary
 * reporter). Exits non-zero, listing every shortfall, when a threshold is
 * not met.
 *
 * Why not jest's built-in coverageThreshold: its global group re-instruments
 * files matched by collectCoverageFrom that it believes lack coverage data
 * — WITHOUT applying coveragePathIgnorePatterns — so the vendored
 * diff-match-patch copy is counted at 0% and deflates the global numbers by
 * ~15 points no matter how coverage is configured. The json-summary report
 * applies the same collection rules the text table shows, so the gates here
 * are calibrated against truthful numbers. Ratchet: only raise these.
 *
 * Current baseline (vendor/scripts excluded): global 73.4/68.0/72.9/74.4,
 * src/lib 90.9/82.8/95.8/93.2 (statements/branches/functions/lines).
 */
const fs = require('fs');
const path = require('path');

const THRESHOLDS = [
  {
    name: 'global',
    match: () => true,
    min: { statements: 70, branches: 65, functions: 70, lines: 72 },
  },
  {
    name: 'src/lib/',
    match: (file) => file.replace(/\\/g, '/').includes('/src/lib/'),
    min: { statements: 88, branches: 80, functions: 93, lines: 91 },
  },
];

const summaryPath = path.join(__dirname, '..', 'coverage', 'coverage-summary.json');
const summary = JSON.parse(fs.readFileSync(summaryPath, 'utf8'));

const failures = [];
for (const group of THRESHOLDS) {
  const metrics = {};
  for (const metric of ['statements', 'branches', 'functions', 'lines']) {
    metrics[metric] = { covered: 0, total: 0 };
  }
  let files = 0;
  for (const [file, data] of Object.entries(summary)) {
    if (file === 'total' || !group.match(file)) continue;
    files++;
    for (const metric of Object.keys(metrics)) {
      metrics[metric].covered += data[metric].covered;
      metrics[metric].total += data[metric].total;
    }
  }
  if (files === 0) {
    failures.push(`group "${group.name}": no files matched — is coverage/coverage-summary.json present?`);
    continue;
  }
  for (const [metric, min] of Object.entries(group.min)) {
    const { covered, total } = metrics[metric];
    const pct = total > 0 ? (covered / total) * 100 : 0;
    if (pct < min) {
      failures.push(
        `group "${group.name}" ${metric}: ${pct.toFixed(2)}% < required ${min}%`
      );
    }
  }
}

if (failures.length > 0) {
  console.error('Coverage gates not met:');
  for (const failure of failures) console.error(`  - ${failure}`);
  process.exit(1);
}
console.log('Coverage gates met.');
