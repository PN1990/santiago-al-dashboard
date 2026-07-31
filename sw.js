// Service worker do PNStays Dashboard.
// - Notificações push (check-ins)
// - Atualização automática: network-first no HTML, para que todos os
//   dispositivos (incluindo iPhone em modo app) recebam sempre a versão
//   mais recente quando estão online, caindo no cache só quando offline.

const CACHE = 'pnstays-v1';
const APP_SHELL = './index.html';

// ── Ciclo de vida: ativa já e limpa caches antigas ────────────────────────────
self.addEventListener('install', function(event) {
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil((async function() {
    const keys = await caches.keys();
    await Promise.all(keys.filter(function(k) { return k !== CACHE; })
                          .map(function(k) { return caches.delete(k); }));
    await self.clients.claim();
  })());
});

// ── Fetch: só intercepta a navegação (o documento HTML) ───────────────────────
// Tudo o resto (Chart.js, SheetJS, chamadas à API noutra origem) passa direto.
self.addEventListener('fetch', function(event) {
  const req = event.request;
  const aceita = req.headers.get('accept') || '';
  const ehNavegacao = req.mode === 'navigate' ||
                      req.destination === 'document' ||
                      aceita.includes('text/html');
  if (req.method !== 'GET' || !ehNavegacao) return;

  event.respondWith((async function() {
    try {
      // network-first: busca sempre a versão fresca, sem cache HTTP
      const fresh = await fetch(req, { cache: 'no-store' });
      const cache = await caches.open(CACHE);
      cache.put(APP_SHELL, fresh.clone());
      return fresh;
    } catch (err) {
      // offline: serve a última versão guardada
      const cache = await caches.open(CACHE);
      const cached = await cache.match(APP_SHELL);
      return cached || Response.error();
    }
  })());
});

// ── Notificações push ─────────────────────────────────────────────────────────
self.addEventListener('push', function(event) {
  let data = { title: '🏡 Santiago AL', body: 'Check-in amanhã!' };
  try {
    if (event.data) data = JSON.parse(event.data.text());
  } catch(e) {}

  event.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: '/icon.png',
      badge: '/icon.png',
      tag: 'checkin-' + (data.reserva_id || Date.now()),
      requireInteraction: true,
      vibrate: [200, 100, 200],
      actions: [
        { action: 'open', title: 'Ver Dashboard' },
        { action: 'dismiss', title: 'Fechar' }
      ]
    })
  );
});

self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  if (event.action !== 'dismiss') {
    event.waitUntil(clients.openWindow('/'));
  }
});
