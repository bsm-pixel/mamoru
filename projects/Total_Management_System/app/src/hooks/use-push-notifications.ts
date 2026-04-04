/**
 * 실시간 푸시 알림 훅
 * Supabase Realtime으로 push_notifications 테이블 구독
 * 새 레코드 INSERT 감지 → 브라우저 Notification + 알림음
 *
 * Server Key / Firebase Admin SDK 불필요
 * 설정의 notifications.sound_enabled와 연동
 */

import { useEffect, useRef } from 'react';
import { useSetting } from './use-settings';
import { createClient } from '@/lib/supabase/client';

export function usePushNotifications() {
  const soundEnabled = useSetting<boolean>('notifications.sound_enabled', false);
  const subscribed = useRef(false);
  const audioRef = useRef<HTMLAudioElement | null>(null);

  useEffect(() => {
    if (!soundEnabled || subscribed.current) return;
    if (typeof window === 'undefined') return;

    // 브라우저 알림 권한 요청
    if ('Notification' in window && Notification.permission === 'default') {
      Notification.requestPermission();
    }

    // 알림음 프리로드
    audioRef.current = new Audio('/notification.wav');
    audioRef.current.volume = 0.5;

    // Supabase Realtime 구독
    const supabase = createClient();
    const channel = supabase
      .channel('push-notifications')
      .on(
        'postgres_changes',
        { event: 'INSERT', schema: 'public', table: 'push_notifications' },
        (payload) => {
          const data = payload.new as { title: string; body: string; url?: string };

          // 브라우저 Notification (백그라운드에서도 표시)
          if ('Notification' in window && Notification.permission === 'granted') {
            const notif = new Notification(data.title, {
              body: data.body,
              icon: '/icon-192.png',
              tag: 'mamoru-push',
              requireInteraction: true,
            });
            notif.onclick = () => {
              window.focus();
              if (data.url) window.location.href = data.url;
              notif.close();
            };
          }

          // 알림음 재생
          if (audioRef.current) {
            audioRef.current.currentTime = 0;
            audioRef.current.play().catch(() => {});
          }
        }
      )
      .subscribe();

    subscribed.current = true;
    console.log('[Push] Realtime 구독 시작');

    return () => {
      supabase.removeChannel(channel);
      subscribed.current = false;
    };
  }, [soundEnabled]);
}
