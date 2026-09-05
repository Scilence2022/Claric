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

const METRICS = ['statements', 'branches', 'functions', 'lines'];
const summaryPath = path.join(__dirname, '..', 'coverage', 'coverage-summary.json');

function isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function checkCoverage(summary, thresholds = THRESHOLDS) {
  const failures = [];
  if (!isObject(summary) || !isObject(summary.total)) {
    return ['summary must be an object with a total object'];
  }
  if (!Array.isArray(thresholds) || thresholds.length === 0) {
    return ['thresholds must be a non-empty array'];
  }
  for (const [index, group] of thresholds.entries()) {
    if (!isObject(group) || typeof group.name !== 'string' || !group.name.trim() ||
        typeof group.match !== 'function' || !isObject(group.min)) {
      failures.push(`threshold group ${index}: requires name, match function, and min object`);
      continue;
    }
    if (Object.keys(group.min).some((metric) => !METRICS.includes(metric))) {
      failures.push(`group "${group.name}": unknown threshold metric`);
    }
    for (const metric of METRICS) {
      const min = group.min[metric];
      if (!Number.isFinite(min) || min < 0 || min > 100) {
        failures.push(`group "${group.name}" ${metric}: threshold must be a finite number between 0 and 100`);
      }
    }
  }
  for (const [file, data] of Object.entries(summary)) {
    for (const metric of METRICS) {
      const value = isObject(data) && data[metric];
      if (!isObject(value) || !Number.isFinite(value.total) || !Number.isFinite(value.covered) ||
          value.total < 0 || value.covered < 0 || value.covered > value.total) {
        failures.push(`file "${file}" ${metric}: requires finite counts with 0 <= covered <= total`);
      }
    }
  }
  if (failures.length > 0) return failures;

  for (const group of thresholds) {
    const metrics = Object.fromEntries(METRICS.map((metric) => [metric, { covered: 0, total: 0 }]));
    let files = 0;
    for (const [file, data] of Object.entries(summary)) {
      if (file === 'total' || !group.match(file)) continue;
      files++;
      for (const metric of METRICS) {
        metrics[metric].covered += data[metric].covered;
        metrics[metric].total += data[metric].total;
      }
    }
    if (files === 0) {
      failures.push(`group "${group.name}": no files matched`);
      continue;
    }
    for (const metric of METRICS) {
      const { covered, total } = metrics[metric];
      if (!Number.isFinite(total) || !Number.isFinite(covered) || total <= 0 || covered < 0 || covered > total) {
        failures.push(`group "${group.name}" ${metric}: requires finite counts with total > 0 and 0 <= covered <= total`);
        continue;
      }
      const pct = (covered / total) * 100;
      const min = group.min[metric];
      if (pct < min) {
        failures.push(`group "${group.name}" ${metric}: ${pct.toFixed(2)}% < required ${min}%`);
      }
    }
  }
  return failures;
}

function main() {
  let failures;
  try {
    failures = checkCoverage(JSON.parse(fs.readFileSync(summaryPath, 'utf8')));
  } catch (error) {
    failures = [`cannot read coverage summary JSON: ${error.message}`];
  }
  if (failures.length > 0) {
    console.error('Coverage gates not met:');
    for (const failure of failures) console.error(`  - ${failure}`);
    process.exitCode = 1;
    return;
  }
  console.log('Coverage gates met.');
}

if (require.main === module) main();

module.exports = { checkCoverage };
