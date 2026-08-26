const js = require('@eslint/js');
const globals = require('globals');

/**
 * ESLint flat config (ESLint 9).
 *
 * File groups with distinct environments:
 *   - src/            browser/Office.js add-in code (ES modules)
 *   - tests/          Jest specs (jsdom environment, transpiled to CJS by babel-jest)
 *   - tests/__mocks__ Jest module mocks written in CommonJS
 *   - scripts/, *.cjs Node.js CommonJS build/deploy tooling
 */
module.exports = [
  {
    ignores: [
      'dist/',
      'node_modules/',
      'coverage/',
      'logs/',
      'assets/',
      'docs/',
      // Vendored third-party library (verbatim upstream copy, CJS) — not
      // held to this repo's lint rules.
      'src/lib/vendor/',
    ],
  },
  js.configs.recommended,
  {
    files: ['src/**/*.js'],
    languageOptions: {
      ecmaVersion: 2023,
      sourceType: 'module',
      globals: {
        ...globals.browser,
        // Office.js injects these globals at runtime
        Office: 'readonly',
        Word: 'readonly',
        OfficeExtension: 'readonly',
        // webpack DefinePlugin replaces process.env.DEFAULT_* at build time
        process: 'readonly',
      },
    },
    rules: {
      // The add-in logs to console deliberately (mirrors the activity log UI);
      // keep it allowed rather than spraying warnings across every module.
      'no-console': 'off',
      'no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
    },
  },
  {
    files: ['tests/**/*.js'],
    languageOptions: {
      ecmaVersion: 2023,
      sourceType: 'module',
      globals: {
        ...globals.browser,
        ...globals.jest,
        // babel-jest transpiles ESM imports to CommonJS, so specs may use
        // require()/module.exports and the Node `global` object directly.
        ...globals.node,
        Word: 'writable',
        Office: 'writable',
      },
    },
    rules: {
      'no-console': 'off',
      // Jest mock implementations routinely accept parameters they ignore
      // (they must match the real signature), so skip arg/caught-binding checks.
      'no-unused-vars': [
        'error',
        { args: 'none', caughtErrors: 'none' },
      ],
    },
  },
  {
    files: ['tests/__mocks__/**/*.js'],
    languageOptions: {
      ecmaVersion: 2023,
      sourceType: 'commonjs',
      globals: {
        ...globals.node,
      },
    },
  },
  {
    files: ['scripts/**/*.cjs', 'webpack.config.cjs', 'jest.config.cjs', 'eslint.config.cjs'],
    languageOptions: {
      ecmaVersion: 2023,
      sourceType: 'commonjs',
      globals: {
        ...globals.node,
      },
    },
    rules: {
      'no-console': 'off',
    },
  },
];
