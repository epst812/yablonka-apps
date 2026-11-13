// Настройки приложения. Эти значения не секретные — их можно хранить в открытом репозитории.
self.APP_CONFIG = {
  // URL веб-приложения Apps Script (…/exec). Пусто = демо-режим.
  API_URL: '',

  // Firebase → Project settings → General → Your apps → Web app → SDK setup and configuration (Config).
  FIREBASE: {
    apiKey: '',
    authDomain: '',
    projectId: '',
    messagingSenderId: '',
    appId: ''
  },

  // Firebase → Project settings → Cloud Messaging → Web Push certificates → Key pair.
  VAPID_KEY: ''
};
