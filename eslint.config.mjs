import {FlatCompat} from '@eslint/eslintrc';
import globals from 'globals';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const compat = new FlatCompat({
  baseDirectory: __dirname,
});

export default [
  {
      ignores: [
        '.github/**',
        'docs/**',
        'node_modules/**',
        'Sidebar.html',
      ],
  },
  ...compat.extends('google').map((config) => ({
    ...config,
    files: ['Code.gs'],
  })),
  {
    files: ['Code.gs'],
    languageOptions: {
      ecmaVersion: 2021,
      sourceType: 'script',
      globals: {
        ...globals.es2021,
        DocumentApp: 'readonly',
        DriveApp: 'readonly',
        HtmlService: 'readonly',
        LockService: 'readonly',
        PropertiesService: 'readonly',
        Session: 'readonly',
        ScriptApp: 'readonly',
        Utilities: 'readonly',
      },
    },
    rules: {
      'max-len': [
        'error',
        {
          code: 100,
          ignoreStrings: true,
          ignoreTemplateLiterals: true,
          ignoreUrls: true,
        },
      ],
      'new-cap': [
        'error',
        {
          newIsCap: true,
          capIsNew: false,
          properties: false,
        },
      ],
      'no-unused-vars': 'off',
      'require-jsdoc': 'off',
      'valid-jsdoc': 'off',
    },
  },
];
