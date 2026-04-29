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
    'src/lib/**/*.js'
  ],
  coverageDirectory: 'coverage',
  verbose: true
};

