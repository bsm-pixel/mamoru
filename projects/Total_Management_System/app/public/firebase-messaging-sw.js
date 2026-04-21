/* eslint-disable no-undef */
// Firebase Messaging Service Worker — 백그라운드 푸시 알림 수신

importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-messaging-compat.js');

// Firebase config는 env에서 가져올 수 없으므로 fetch로 로드
// 또는 빌드 시 삽입 — 여기서는 self.__FIREBASE_CONFIG 사용
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));

// 메시지 수신 시 알림 표시 (백그라운드)
self.addEventListener('push', (event) => {
  if (!event.data) return;

  let payload;
  try {
    payload = event.data.json();
  } catch {
    payload = { notification: { title: 'MAMORU', body: event.data.text() } };
  }

  const notif = payload.notification || {};
  const title = notif.title || 'MAMORU TMS';
  const options = {
    body: notif.body || '',
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    tag: payload.data?.tag || 'mamoru-default',
    data: payload.data || {},
    requireInteraction: true,
  };

  event.waitUntil(self.registration.showNotification(title, options));
});

// 클라이언트에서 특정 tag 알림 회수 요청 수신
// { type: 'DISMISS', tag: 'mamoru-as_received-AS-YYYYMMDD-NNN' }
self.addEventListener('message', (event) => {
  const data = event.data || {};
  if (data.type === 'DISMISS' && data.tag) {
    event.waitUntil(
      self.registration.getNotifications({ tag: data.tag }).then((notifs) => {
        notifs.forEach((n) => n.close());
      })
    );
  }
});

// 알림 클릭 시 TMS 열기
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const url = event.notification.data?.url || '/dashboard';

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clients) => {
      // 이미 열린 TMS 탭이 있으면 포커스
      for (const client of clients) {
        if (client.url.includes('/dashboard') || client.url.includes('/consultations') || client.url.includes('/repairs')) {
          client.focus();
          if (url !== '/dashboard') client.navigate(url);
          return;
        }
      }
      // 없으면 새 탭
      return self.clients.openWindow(url);
    })
  );
});
