/* eslint-disable no-undef */
// Firebase Messaging Service Worker — 백그라운드 푸시 알림 수신
//
// ⚠️ 외부 스크립트(gstatic firebase compat) 로드하지 않는다 — push 핸들러가
//    event.data.json() 파싱 + showNotification 만 쓰므로 firebase 라이브러리 불필요.
//    (느린/끊긴 회선서 importScripts 실패 → SW 설치 실패 → 알림 미수신 이던 문제 제거, 2026-08-01)
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
  // tag 우선순위: data.tag → notification.tag (webpush.notification 경로) → fallback
  // Realtime/window.Notification 경로와 동일 tag 사용 → 브라우저 자동 dedup
  const options = {
    body: notif.body || '',
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    tag: payload.data?.tag || notif.tag || 'mamoru-default',
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
// 우선순위: ① 이미 열린 same-origin TMS 인스턴스(PWA/탭)가 있으면 그걸 focus + navigate
//           ② 없을 때만 새 창 생성
// 어느 페이지에 있든 (sales/orders/customers/manual-invoices 등) 매칭되도록 origin 기준으로 매칭
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const targetUrl = event.notification.data?.url || '/dashboard';

  event.waitUntil((async () => {
    const allClients = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });

    // 같은 origin (TMS) 클라이언트 후보 — PWA standalone 인스턴스도 여기 포함됨
    const candidates = allClients.filter((c) => {
      try {
        return new URL(c.url).origin === self.location.origin;
      } catch {
        return false;
      }
    });

    if (candidates.length > 0) {
      // 가장 최근 활성화된 것이 일반적으로 배열 앞쪽이지만, focused 우선 → 그 외 첫 항목
      const focused = candidates.find((c) => c.focused) || candidates[0];

      // PWA standalone(Android Chrome 등)에서 focused.navigate() 가 막히는 경우 대비:
      // 클라이언트(NotificationNavigator)에 postMessage 로 navigate 요청 → Next.js router 가 처리 (2026-05-23 fix)
      try {
        focused.postMessage({ type: 'NAVIGATE_FROM_NOTIFICATION', url: targetUrl });
      } catch {
        // postMessage 실패해도 navigate 시도는 진행
      }

      try {
        if ('navigate' in focused && typeof focused.navigate === 'function') {
          await focused.navigate(targetUrl);
        }
      } catch {
        // navigate 막힘 → postMessage 가 클라이언트에서 처리할 것
      }
      return focused.focus();
    }

    // 같은 origin 인스턴스가 전혀 없을 때만 새 창
    return self.clients.openWindow(targetUrl);
  })());
});
