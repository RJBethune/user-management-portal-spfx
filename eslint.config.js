const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');
const { plugin: shadcn } = require('@shadcn/lint');

module.exports = [
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    plugins: { shadcn },
    settings: { shadcn: { componentImports: ['^@spfx-kit/ui-profile(/|$)', '^@spfx-kit/ui-patterns(/|$)'] } },
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    },
    rules: {
      '@typescript-eslint/no-floating-promises': 'off',
      'shadcn/no-restyle': ['error', { allow: ['layout'] }],
      'shadcn/no-raw-colors': 'error',
      'shadcn/no-arbitrary-values': 'error'
    }
  }
];
