/**
 * 푸시 알림 구독 훅
 * 로그인 후 자동으로 FCM 토큰을 발급받아 DB에 저장
 */

import { useEffect, useRef } from 'react';
import { useSetting } from './use-settings';

export function usePushNotifications() {
  const soundEnabled = useSetting<boolean>('notifications.sound_enabled', false);
  const subscribed = useRef(false);

  useEffect(() => {
    if (!soundEnabled || subscribed.current) return;
    if (typeof window === 'undefined' || !('Notification' in window)) return;

    // Firebase 환경변수 확인
    if (!process.env.NEXT_PUBLIC_FIREBASE_API_KEY) return;

    (async () => {
      try {
        const { requestPushToken } = await import('@/lib/firebase/client');
        const token = await requestPushToken();
        if (!token) return;

        // 서버에 토큰 등록
        await fetch('/api/push/subscribe', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            token,
            deviceInfo: `${navigator.userAgent.slice(0, 100)}`,
          }),
        });

        subscribed.current = true;
        console.log('[Push] 구독 완료');
      } catch (err) {
        console.error('[Push] 구독 실패:', err);
      }
    })();
  }, [soundEnabled]);
}
