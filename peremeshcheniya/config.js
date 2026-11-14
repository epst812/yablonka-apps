// Настройки приложения. Эти значения не секретные — их можно хранить в открытом репозитории.
self.APP_CONFIG = {
  // URL веб-приложения Apps Script (…/exec). Пусто = демо-режим.
  API_URL: 'https://script.google.com/macros/s/AKfycbxEG4SO5PJs7OSV8xEZpbf-_Akm9Z4vurznSq-Ccm4uIKtePhzgCVrfHDWOPApY8XCI/exec',

  // Firebase → Project settings → General → Your apps → Web app → SDK setup and configuration (Config).
  FIREBASE: {
    apiKey: 'AIzaSyC0lt69vwND6bdgF-NXia0ug6Q1yMb158g',
    authDomain: 'yablonka-peremeshcheniya.firebaseapp.com',
    projectId: 'yablonka-peremeshcheniya',
    messagingSenderId: '1023249158044',
    appId: '1:1023249158044:web:bbb08b276ab3b4c0c0c259'
  },

  // Firebase → Project settings → Cloud Messaging → Web Push certificates → Key pair.
  VAPID_KEY: 'BOk1VZu0dDgx8l-aZMS6MUKBqRIl7jPK8gLXsCGucGyEtLIYb8Dwm69X8z58Hs5h0zBOp--Cm_YhsunfUkSJ6UM'
};
