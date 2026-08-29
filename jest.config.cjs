// Test environment strategy: 'node' by default. Specs that need a DOM
// either declare `@jest-environment jsdom` in a docblock (6 specs) or
// construct JSDOM manually via `new JSDOM(...)` (3 specs).
module.exports = {
  testEnvironment: 'node',
  transform: {
    '^.+\\.js$': 'babel-jest'
  },
  transformIgnorePatterns: [
    'node_modules/(?!marked/)'
  ],
  moduleNameMapper: {
    // Stub CSS imports so taskpane.js (which does `import './taskpane.css'`
    // for webpack) can be required from Jest specs without a CSS parser.
    '\\.css$': '<rootDir>/tests/__mocks__/styleMock.js'
  },
  moduleFileExtensions: ['js'],
  testMatch: ['**/tests/**/*.spec.js'],
  collectCoverageFrom: [
    'src/**/*.js'
  ],
  // The vendored diff-match-patch copy and the one-off utility scripts are
  // not held to the coverage gates. (Note: jest's built-in
  // coverageThreshold re-instruments collectCoverageFrom matches it deems
  // uncovered WITHOUT applying these ignore patterns, so the vendored file
  // is counted at 0% and deflates the global numbers ~15 points — the gate
  // therefore lives in scripts/check-coverage.cjs over the json-summary
  // report, which matches the table.)
  coveragePathIgnorePatterns: [
    '/node_modules/',
    '<rootDir>/src/lib/vendor/',
    '<rootDir>/src/scripts/'
  ],
  // NOTE: no coverageThreshold here on purpose — jest's threshold check
  // re-instruments collectCoverageFrom matches it deems uncovered WITHOUT
  // applying the ignore patterns above, counting the vendored diff-match-
  // patch copy at 0% and deflating the global numbers ~15 points. The gates
  // live in scripts/check-coverage.cjs over the json-summary report, whose
  // numbers match the table: global 70/65/70/72, src/lib 88/80/93/91
  // (baseline 73.4/68.0/72.9/74.4 and 90.9/82.8/95.8/93.2 — ratchet, only
  // raise).
  coverageDirectory: 'coverage',
  verbose: true
};

