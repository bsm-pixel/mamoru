/**
 * Firebase Admin — 서버에서 푸시 알림 발송
 * API route에서 호출
 */

const FCM_SERVER_KEY = (process.env.FIREBASE_SERVER_KEY || '').trim();

interface PushPayload {
  title: string;
  body: string;
  url?: string;
  tag?: string;
}

/** 등록된 모든 디바이스에 푸시 발송 */
export async function sendPushToAll(payload: PushPayload): Promise<{ sent: number; failed: number }> {
  if (!FCM_SERVER_KEY) {
    console.warn('[FCM] FIREBASE_SERVER_KEY 미설정 — 푸시 발송 스킵');
    return { sent: 0, failed: 0 };
  }

  // DB에서 구독 토큰 조회
  const { createServiceClient } = await import('@/lib/supabase/server');
  const db = createServiceClient();
  const { data: subs } = await (db as any).from('push_subscriptions').select('token'); // eslint-disable-line @typescript-eslint/no-explicit-any

  if (!subs || subs.length === 0) {
    console.log('[FCM] 구독 토큰 없음');
    return { sent: 0, failed: 0 };
  }

  let sent = 0;
  let failed = 0;

  for (const sub of subs) {
    try {
      const res = await fetch('https://fcm.googleapis.com/fcm/send', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `key=${FCM_SERVER_KEY}`,
        },
        body: JSON.stringify({
          to: sub.token,
          notification: {
            title: payload.title,
            body: payload.body,
            icon: '/icon-192.png',
            click_action: payload.url || '/dashboard',
          },
          data: {
            url: payload.url || '/dashboard',
            tag: payload.tag || 'mamoru',
          },
        }),
      });

      if (res.ok) {
        sent++;
      } else {
        failed++;
        const err = await res.text();
        console.error('[FCM] 발송 실패:', err);
        // 토큰 만료 시 삭제
        if (err.includes('NotRegistered') || err.includes('InvalidRegistration')) {
          await (db as any).from('push_subscriptions').delete().eq('token', sub.token); // eslint-disable-line @typescript-eslint/no-explicit-any
        }
      }
    } catch (e) {
      failed++;
      console.error('[FCM] 발송 에러:', e);
    }
  }

  console.log(`[FCM] 발송 완료: ${sent}건 성공, ${failed}건 실패`);
  return { sent, failed };
}
