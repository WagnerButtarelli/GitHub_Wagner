self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(clients.claim()));
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const msg = { type: 'act', id: e.notification.data.id, action: e.action || 'open' };
  e.waitUntil(clients.matchAll({ type: 'window', includeUncontrolled: true }).then(cs => {
    cs.forEach(c => c.postMessage(msg));
    if (!cs.length) return clients.openWindow('./bridge.html');
    if (msg.action === 'open') return cs[0].focus();
  }));
});
