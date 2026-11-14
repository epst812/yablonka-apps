// Service worker: установка приложения, офлайн-кэш страницы и пуш-уведомления.
importScripts('config.js');
const CONF = self.APP_CONFIG || {};
const FB_VER = '10.12.2';
const PUSH_ON = !!(CONF.FIREBASE && CONF.FIREBASE.apiKey);
if (PUSH_ON) {
  importScripts(
    'https://www.gstatic.com/firebasejs/' + FB_VER + '/firebase-app-compat.js',
    'https://www.gstatic.com/firebasejs/' + FB_VER + '/firebase-messaging-compat.js'
  );
  firebase.initializeApp(CONF.FIREBASE);
  // Сервер шлёт data-сообщения (title, body, id) — показываем их сами.
  firebase.messaging().onBackgroundMessage(async msg => {
    const d = msg.data || {};
    // если приложение открыто где-то в фоне — пусть сразу подтянет новую заявку
    const wins = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    wins.forEach(w => w.postMessage({ type: 'refresh', id: d.id || '' }));
    return self.registration.showNotification(d.title || 'Перемещения', {
      body: d.body || '',
      tag: d.tag || 'peremeshcheniya',
      renotify: true,
      icon: 'icon-192.png',
      badge: 'icon-192.png',
      data: { id: d.id || '' }
    });
  });
}

self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil((async () => {
    const wins = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    const id = (e.notification.data && e.notification.data.id) || '';
    for (const w of wins) { if ('focus' in w) { w.postMessage({ type: 'refresh', id }); return w.focus(); } }
    return clients.openWindow('./' + (id ? '?open=' + encodeURIComponent(id) : ''));
  })());
});

const CACHE = 'peremeshcheniya-v5';
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', 'index.html', 'manifest.json', 'icon-192.png', 'logo.png'])));
  self.skipWaiting();
});
self.addEventListener('activate', e => e.waitUntil((async () => {
  for (const k of await caches.keys()) if (k !== CACHE) await caches.delete(k);
  await self.clients.claim();
})()));
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // Google и Firebase не трогаем
  e.respondWith(
    fetch(new Request(e.request.url, { cache: 'no-cache' })).then(r => {
      const copy = r.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return r;
    }).catch(() => caches.match(e.request))
  );
});
