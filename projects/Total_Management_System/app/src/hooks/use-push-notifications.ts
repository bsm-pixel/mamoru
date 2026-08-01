/**
 * 실시간 푸시 알림 훅
 *
 * 1. FCM 토큰 발급 → 서버 저장 (모바일 백그라운드 푸시)
 * 2. Supabase Realtime 구독 (포그라운드 알림음 + 브라우저 Notification)
 */

import { useEffect, useRef } from 'react';
import { useSetting } from './use-settings';
import { createClient } from '@/lib/supabase/client';

export function usePushNotifications() {
  const soundEnabled = useSetting<boolean>('notifications.sound_enabled', true);
  const subscribed = useRef(false);
  const fcmRegistered = useRef(false);
  const audioRef = useRef<HTMLAudioElement | null>(null);

  // FCM 토큰 발급 + 서버 등록 (1회)
  useEffect(() => {
    if (fcmRegistered.current) return;
    if (typeof window === 'undefined') return;
    if (!('serviceWorker' in navigator)) return;

    fcmRegistered.current = true;

    (async () => {
      try {
        const { requestPushToken } = await import('@/lib/firebase/client');
        const token = await requestPushToken();
        if (!token) return;

        // 서버에 토큰 저장
        await fetch('/api/push/subscribe', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ token }),
        });
        console.log('[Push] FCM 토큰 등록 완료');
      } catch (err) {
        console.error('[Push] FCM 토큰 등록 실패:', err);
      }
    })();
  }, []);

  // Supabase Realtime 구독 (포그라운드 알림)
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
          const data = payload.new as { title: string; body: string; url?: string; tag?: string };

          // 브라우저 Notification
          // tag를 data.tag로 사용 → FCM Service Worker와 동일 tag → 브라우저 자동 dedup (중복 표시 방지)
          if ('Notification' in window && Notification.permission === 'granted') {
            const notif = new Notification(data.title, {
              body: data.body,
              icon: '/icon-192.png',
              tag: data.tag || 'mamoru-push',
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
