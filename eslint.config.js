const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  // Build output - never lint generated bundles.
  { ignores: ['sharepoint/**', 'lib/**', 'lib-*/**', 'dist/**', 'temp/**', 'release/**', 'jest-output/**'] },
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    }
  }
];
