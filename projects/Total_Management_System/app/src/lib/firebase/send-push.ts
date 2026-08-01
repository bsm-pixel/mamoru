/**
 * Firebase Admin SDK — 서버에서 푸시 알림 발송
 * 모바일 백그라운드에서도 알림 수신 가능 (FCM)
 */

import * as admin from 'firebase-admin';

// Firebase Admin 초기화 (싱글톤)
function getApp() {
  if (admin.apps.length > 0) return admin.apps[0]!;

  const projectId = (process.env.FIREBASE_PROJECT_ID || '').trim();
  const clientEmail = (process.env.FIREBASE_CLIENT_EMAIL || '').trim();
  const privateKey = (process.env.FIREBASE_PRIVATE_KEY || '').trim().replace(/\\n/g, '\n');

  if (!projectId || !clientEmail || !privateKey) {
    console.warn('[FCM] Firebase Admin 환경변수 미설정 — 푸시 비활성');
    return null;
  }

  return admin.initializeApp({
    credential: admin.credential.cert({ projectId, clientEmail, privateKey }),
  });
}

interface PushPayload {
  title: string;
  body: string;
  url?: string;
  tag?: string;
  /** @deprecated 설정 게이팅 제거됨(2026-08-01) — 고객 행동 푸시는 항상 발송. 값은 무시된다. */
  settingKey?: string;
}

/**
 * 등록된 모든 디바이스에 FCM 푸시 발송 (무조건 발송 — on/off 게이팅 없음).
 * 🔴 고객 접수/행동 알림은 사장님이 놓치면 안 되므로 어떤 설정으로도 차단하지 않는다.
 *    (2026-08-01: isPushEnabled 게이팅 제거 — 접수 알림 오락가락 근본원인 정리)
 * 발송은 sendEach 배치 1회로 처리 → 순차 send 루프의 지연/부분누락 제거.
 */
export async function sendPushToAll(payload: PushPayload): Promise<{ sent: number; failed: number }> {
  const app = getApp();
  if (!app) return { sent: 0, failed: 0 };

  const { createServiceClient } = await import('@/lib/supabase/server');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;

  // Realtime 폴백 저장 (TMS 탭 열려있으면 소리) — 토큰 유무와 무관하게 항상
  const saveRealtime = () =>
    db.from('push_notifications').insert({
      title: payload.title,
      body: payload.body,
      url: payload.url || '/dashboard',
      tag: payload.tag || 'mamoru',
      read: false,
    });

  const { data: subs } = await db.from('push_subscriptions').select('token');
  if (!subs || subs.length === 0) {
    console.log('[FCM] 구독 토큰 없음');
    await saveRealtime();
    return { sent: 0, failed: 0 };
  }

  const messaging = admin.messaging(app);
  const tokens: string[] = subs.map((s: { token: string }) => s.token);
  const messages = tokens.map((token) => ({
    token,
    notification: { title: payload.title, body: payload.body },
    webpush: {
      notification: {
        icon: '/icon-192.png',
        badge: '/icon-192.png',
        tag: payload.tag || 'mamoru',
        requireInteraction: true,
      },
      fcmOptions: { link: payload.url || '/dashboard' },
    },
    data: {
      url: payload.url || '/dashboard',
      tag: payload.tag || 'mamoru', // SW/Realtime 중복 dedup 용
    },
  }));

  let sent = 0;
  let failed = 0;
  try {
    const resp = await messaging.sendEach(messages);
    sent = resp.successCount;
    failed = resp.failureCount;
    // 만료/무효 토큰만 정리
    const stale: string[] = [];
    resp.responses.forEach((r, i) => {
      if (!r.success) {
        const code = r.error?.code || '';
        console.error('[FCM] 발송 실패:', code, r.error?.message);
        if (code.includes('not-registered') || code.includes('invalid-registration') || code.includes('invalid-argument')) {
          stale.push(tokens[i]);
        }
      }
    });
    if (stale.length) await db.from('push_subscriptions').delete().in('token', stale);
  } catch (err) {
    failed = messages.length;
    console.error('[FCM] sendEach 오류:', err instanceof Error ? err.message : String(err));
  }

  await saveRealtime();
  console.log(`[FCM] 발송: ${sent}건 성공, ${failed}건 실패`);
  return { sent, failed };
}
