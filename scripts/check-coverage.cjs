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
 * Current baseline (vendor/scripts excluded), statements/branches/functions/lines:
 *   global        73.6/68.6/73.8/74.6
 *   src/lib/      90.5/82.4/94.4/92.4
 *   src/taskpane/ 57.2/52.8/55.6/57.7
 *
 * Every group carries its own gate on purpose. With only `global` and
 * `src/lib/`, the taskpane layer was unconstrained: src/lib is the larger
 * share of the codebase, so its ~90% held `global` above 70 while
 * src/taskpane could fall to ~50% with CI still green. The gate was tightest
 * exactly where discipline was already best, and absent where it was needed.
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
  {
    // Set just below the measured baseline (57.2/52.8/55.6/57.7) so normal
    // variation does not fail CI, while a real regression does. Ratchet these
    // up as the UI layer gains tests — settings-view.js (733 statements, 0%)
    // and status-bar.js (97 statements, 22%) are the largest gaps.
    name: 'src/taskpane/',
    match: (file) => file.replace(/\\/g, '/').includes('/src/taskpane/'),
    min: { statements: 56, branches: 51, functions: 54, lines: 56 },
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
