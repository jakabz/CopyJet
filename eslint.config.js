const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    }
  },
  {
    // The core is UI-independent (CLAUDE.md): no React, Fluent UI or SPFx web part imports.
    files: ['src/core/**/*.ts', 'src/core/**/*.tsx'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          patterns: [
            { group: ['react', 'react-dom', 'react/*', 'react-dom/*'], message: 'src/core must not depend on React.' },
            { group: ['@fluentui/*', 'office-ui-fabric-react', 'office-ui-fabric-react/*'], message: 'src/core must not depend on Fluent UI.' },
            { group: ['@microsoft/sp-*'], message: 'src/core must not depend on SPFx packages; take structural types instead.' },
            { group: ['**/webparts/**', '**/shared/**'], message: 'src/core must not import UI code.' }
          ]
        }
      ]
    }
  }
];
