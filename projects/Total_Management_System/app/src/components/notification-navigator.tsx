'use client';

import { useEffect } from 'react';
import { useRouter } from 'next/navigation';

/**
 * Service Worker → 클라이언트 navigate 위임 (2026-05-23 fix)
 *
 * 배경: Android Chrome PWA standalone 모드에서 SW 의 `focused.navigate(url)` API 가
 *       종종 막혀서, 알림 클릭 시 PWA 가 focus 만 되고 페이지 이동이 안 되는 케이스가 있음.
 *
 * 해결: SW (firebase-messaging-sw.js) 의 notificationclick 핸들러에서
 *       `client.postMessage({ type: 'NAVIGATE_FROM_NOTIFICATION', url })` 를 보내고,
 *       이 컴포넌트(클라이언트)가 받아서 Next.js router.push() 로 이동.
 *
 * root layout 에 한 번 마운트하면 모든 페이지에서 동작.
 */
export function NotificationNavigator() {
  const router = useRouter();

  useEffect(() => {
    if (typeof navigator === 'undefined' || !('serviceWorker' in navigator)) return;

    const handler = (event: MessageEvent) => {
      const data = event.data;
      if (data && data.type === 'NAVIGATE_FROM_NOTIFICATION' && typeof data.url === 'string') {
        // 이미 같은 경로면 무시 (중복 navigate 방지)
        if (window.location.pathname + window.location.search === data.url) return;
        router.push(data.url);
      }
    };

    navigator.serviceWorker.addEventListener('message', handler);
    return () => navigator.serviceWorker.removeEventListener('message', handler);
  }, [router]);

  return null;
}
