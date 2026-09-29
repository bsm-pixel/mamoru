/**
 * 이 기기의 푸시 토큰 발급 + 서버 등록 (2026-09-29 훅에서 분리)
 *
 * 왜 분리했나: 전엔 대시보드 안(로그인 후)에서만 재등록돼서, 토큰이 죽고 로그인까지 풀린 휴대폰은
 * "다시 로그인할 때까지" 알림이 비었다. 이제 **로그인 화면에서도** 재등록한다
 * (서버는 이미 등록된 기기 id 면 로그인 없이도 토큰을 교체해준다 — api/push/subscribe).
 */

import { requestPushToken } from './client';
import { getDeviceId, getDeviceLabel } from './device';

/** SW 가 수신 확인(ack) 때 읽는 기기 id 위치 — firebase-messaging-sw.js readDeviceId() 와 짝 */
const DEVICE_CACHE = 'mamoru-meta';
const DEVICE_CACHE_KEY = '/__mamoru_device';

export async function registerPushDevice(opts: { allowPrompt: boolean }): Promise<void> {
  if (typeof window === 'undefined') return;
  if (!('serviceWorker' in navigator) || !('Notification' in window)) return;
  // 로그인 화면 등에서는 권한 팝업을 띄우지 않는다 — 이미 허용된 기기만 조용히 재등록
  if (!opts.allowPrompt && Notification.permission !== 'granted') return;

  // 저장공간이 부족해도 브라우저가 이 앱 데이터(서비스워커·토큰·기기 id)를 먼저 지우지 않게 요청
  navigator.storage?.persist?.().catch(() => {});

  const token = await requestPushToken();
  if (!token) return;

  const deviceId = getDeviceId();
  const deviceInfo = getDeviceLabel();

  if (deviceId) {
    try {
      const cache = await caches.open(DEVICE_CACHE);
      await cache.put(DEVICE_CACHE_KEY, new Response(deviceId));
    } catch { /* Cache API 없음 — SW 는 userAgent 로 기기 판별 */ }
  }

  // 🔴 2026-09-15: 기기 id를 함께 보낸다 — 같은 기기의 옛 토큰만 교체(다른 기기는 건드리지 않음)
  const res = await fetch('/api/push/subscribe', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ token, deviceId, deviceInfo }),
  });
  console.log('[Push] FCM 토큰 등록', res.ok ? '완료' : `실패(${res.status})`, '—', deviceInfo);
}
