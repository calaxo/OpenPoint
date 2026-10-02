// Lint minimal : attrape les erreurs réelles (variable non définie, import manquant, code mort),
// sans imposer de style. Lancement : npm run lint
import js from '@eslint/js';
import globals from 'globals';

export default [
  { ignores: ['node_modules/', 'renderer/_work/', 'renderer/fonts/', 'binary/', 'test/fixtures/out/', '**/*.html'] },
  js.configs.recommended,
  {
    files: ['**/*.{js,mjs,cjs}'],
    languageOptions: { ecmaVersion: 2024, sourceType: 'module', globals: { ...globals.node } },
    rules: {
      'no-unused-vars': ['warn', { args: 'none', caughtErrors: 'none' }],
      'no-empty': ['error', { allowEmptyCatch: true }],
      'no-control-regex': 'off',
      'no-useless-escape': 'off',
      'no-misleading-character-class': 'off',
    },
  },
  { files: ['renderer/viewer/**/*.js'], languageOptions: { sourceType: 'script', globals: { ...globals.browser } } },
];
