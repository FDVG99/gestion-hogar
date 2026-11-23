const js = require('@eslint/js');

module.exports = [
  {
    ignores: ['node_modules/**', '.git/**'],
  },

  {
    ...js.configs.recommended,
    files: ['apps-script/**/*.js'],
  },

  {
    files: ['apps-script/**/*.js'],

    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'script',

      globals: {
        // Google Apps Script
        SpreadsheetApp: 'readonly',
        HtmlService: 'readonly',
        ScriptApp: 'readonly',
        PropertiesService: 'readonly',
        LockService: 'readonly',
        CacheService: 'readonly',
        DriveApp: 'readonly',
        GmailApp: 'readonly',
        MailApp: 'readonly',
        UrlFetchApp: 'readonly',
        Utilities: 'readonly',
        Session: 'readonly',
        ContentService: 'readonly',
        Logger: 'readonly',
        FormApp: 'readonly',
        CalendarApp: 'readonly',
        DocumentApp: 'readonly',
        console: 'readonly',
      },
    },

    rules: {
      'no-unused-vars': [
        'warn',
        {
          vars: 'local',
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
    },
  },

  {
    files: ['apps-script/Code.js'],

    languageOptions: {
      globals: {
        getAppData: 'readonly',
      },
    },
  },
];
