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
        // 🔴 2026-09-29: 등록 로직은 lib/firebase/register.ts 로 분리 (로그인 화면에서도 재등록)
        const { registerPushDevice } = await import('@/lib/firebase/register');
        await registerPushDevice({ allowPrompt: true });
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

          // 브라우저 알림 — 2026-11-22: 탭이 직접 띄우던 것(new Notification)을 서비스 워커로 옮김.
          //   윈도우 크롬은 탭이 띄운 알림과 서비스 워커(FCM)가 띄운 알림을 다른 출처로 봐서, 같은 tag 여도
          //   하나로 합치지 못하고 두 번 떴다(「각인 문구 도착」 2번). 같은 곳(서비스 워커)에서 같은 tag 로 띄우면
          //   먼저 뜬 것을 조용히 바꿔치기 한다. 누르면 서비스 워커의 notificationclick 이 data.url 로 연다.
          if ('Notification' in window && Notification.permission === 'granted') {
            const tag = data.tag || 'mamoru-push';
            const showInPage = () => {
              const notif = new Notification(data.title, { body: data.body, icon: '/icon-192.png', tag, requireInteraction: true });
              notif.onclick = () => { window.focus(); if (data.url) window.location.href = data.url; notif.close(); };
            };
            const sw = 'serviceWorker' in navigator ? navigator.serviceWorker.getRegistration() : Promise.resolve(undefined);
            sw.then((reg) => {
              if (!reg) { showInPage(); return; }
              return reg.showNotification(data.title, {
                body: data.body, icon: '/icon-192.png', badge: '/icon-192.png', tag,
                data: { url: data.url || '/dashboard', tag }, requireInteraction: true,
              });
            }).catch(showInPage);
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
