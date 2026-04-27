/**
 * Firebase Client — 푸시 알림 구독 (브라우저)
 */

let messaging: any = null; // eslint-disable-line @typescript-eslint/no-explicit-any

/** Firebase 초기화 (클라이언트 사이드만) */
async function getMessaging() {
  if (messaging) return messaging;
  if (typeof window === 'undefined') return null;

  const apiKey = process.env.NEXT_PUBLIC_FIREBASE_API_KEY;
  const projectId = process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID;
  const messagingSenderId = process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID;
  const appId = process.env.NEXT_PUBLIC_FIREBASE_APP_ID;

  if (!apiKey || !projectId || !messagingSenderId || !appId) {
    console.warn('[Firebase] 환경변수 미설정 — 푸시 알림 비활성');
    return null;
  }

  const { initializeApp, getApps } = await import('firebase/app');
  const { getMessaging: getMsg } = await import('firebase/messaging');

  const config = { apiKey, authDomain: `${projectId}.firebaseapp.com`, projectId, messagingSenderId, appId };

  const app = getApps().length === 0 ? initializeApp(config) : getApps()[0];
  messaging = getMsg(app);

  // 포그라운드 메시지 핸들러는 의도적으로 비워둠.
  // Service Worker(/firebase-messaging-sw.js)의 'push' 이벤트가 백그라운드/포그라운드
  // 모두에서 showNotification을 호출하므로, 여기서 또 new Notification을 만들면
  // 같은 푸시가 두 번 표시되는 중복 문제가 발생함.
  // 향후 토스트/사운드 등 inline UI가 필요하면 그때 onMessage 추가.

  return messaging;
}

/** FCM 토큰 발급 (사용자 알림 허용 필요) */
export async function requestPushToken(): Promise<string | null> {
  try {
    const msg = await getMessaging();
    if (!msg) return null;

    // Service Worker 등록
    const registration = await navigator.serviceWorker.register('/firebase-messaging-sw.js');

    const permission = await Notification.requestPermission();
    if (permission !== 'granted') {
      console.log('[Firebase] 알림 권한 거부');
      return null;
    }

    const vapidKey = process.env.NEXT_PUBLIC_FIREBASE_VAPID_KEY;
    if (!vapidKey) {
      console.warn('[Firebase] VAPID_KEY 미설정');
      return null;
    }

    const { getToken } = await import('firebase/messaging');
    const token = await getToken(msg, { vapidKey, serviceWorkerRegistration: registration });

    console.log('[Firebase] FCM 토큰 발급:', token?.slice(0, 20) + '...');
    return token;
  } catch (err) {
    console.error('[Firebase] 토큰 발급 실패:', err);
    return null;
  }
}
