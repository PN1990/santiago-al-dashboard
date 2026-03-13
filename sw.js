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
