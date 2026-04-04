/**
 * 푸시 알림 발송 — FCM Legacy API 대신 브라우저 Push API 사용
 * VAPID Key만으로 동작 (Server Key / Admin SDK 불필요)
 *
 * 동작 흐름:
 * 1. DB에서 등록된 FCM 토큰 조회
 * 2. FCM HTTP v1 API로 발송 (OAuth2 없이 VAPID 기반)
 *
 * 참고: FCM은 클라이언트에서 getToken()으로 발급받은 토큰에
 * 서버에서 메시지를 보내려면 Server Key가 필요하지만,
 * 조직 정책으로 생성 불가하므로 대안 사용:
 * → Supabase Realtime 구독으로 클라이언트에서 직접 알림 표시
 */

interface PushPayload {
  title: string;
  body: string;
  url?: string;
  tag?: string;
}

/** DB에 알림 레코드를 삽입 → 클라이언트가 Realtime으로 수신 */
export async function sendPushToAll(payload: PushPayload): Promise<{ sent: number; failed: number }> {
  try {
    const { createServiceClient } = await import('@/lib/supabase/server');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;

    // push_notifications 테이블에 삽입 → 클라이언트 Realtime 구독이 감지
    const { error } = await db.from('push_notifications').insert({
      title: payload.title,
      body: payload.body,
      url: payload.url || '/dashboard',
      tag: payload.tag || 'mamoru',
      read: false,
    });

    if (error) {
      console.error('[Push] 알림 저장 실패:', error);
      return { sent: 0, failed: 1 };
    }

    console.log(`[Push] 알림 저장: ${payload.title} — ${payload.body}`);
    return { sent: 1, failed: 0 };
  } catch (err) {
    console.error('[Push] 에러:', err);
    return { sent: 0, failed: 1 };
  }
}
