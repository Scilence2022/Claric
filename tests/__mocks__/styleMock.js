// Jest CSS import stub.
// taskpane.js does `import './taskpane.css'` for webpack to bundle styles;
// under Jest (node + jsdom), CSS is irrelevant and would crash the parser.
// Mapped via jest.config.cjs `moduleNameMapper: { '\\.css$': ... }`.
module.exports = {};
