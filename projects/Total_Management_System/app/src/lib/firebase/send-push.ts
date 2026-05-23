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
  /** 설정 키 (system_settings의 push.* 키). 설정 OFF면 발송 안 함 */
  settingKey?: string;
}

/** 푸시 이벤트 설정 키 → 기본값 매핑 (설정값이 없으면 이 값 사용) */
const PUSH_DEFAULTS: Record<string, boolean> = {
  'push.consultation_received': true,
  'push.field_request': true,
  'push.talk_received': true,
  'push.field_confirmed': true,
  'push.field_reschedule': true,
  'push.field_cancelled': true,         // 출장 예약 취소 (2026-05-23 추가)
  'push.consultation_cancelled': true,  // 매장방문·톡상담 예약 취소 (2026-05-23 추가)
  'push.repair_received': true,
  'push.review_submitted': true,
  'push.order_received': true,
};

async function isPushEnabled(settingKey?: string): Promise<boolean> {
  if (!settingKey) return true;
  try {
    const { createServiceClient } = await import('@/lib/supabase/server');
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data } = await (db as any).from('system_settings').select('value').eq('key', settingKey).single();
    if (!data) return PUSH_DEFAULTS[settingKey] ?? true;
    const v = data.value;
    if (v === 'false' || v === false) return false;
    return true;
  } catch {
    return true; // DB 오류 시 발송 (안전)
  }
}

/** 등록된 모든 디바이스에 FCM 푸시 발송 */
export async function sendPushToAll(payload: PushPayload): Promise<{ sent: number; failed: number }> {
  // 설정 기반 on/off 체크
  const enabled = await isPushEnabled(payload.settingKey);
  if (!enabled) {
    console.log(`[FCM] SKIP ${payload.settingKey} — 설정에서 비활성`);
    return { sent: 0, failed: 0 };
  }

  const app = getApp();
  if (!app) return { sent: 0, failed: 0 };

  // DB에서 구독 토큰 조회
  const { createServiceClient } = await import('@/lib/supabase/server');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;
  const { data: subs } = await db.from('push_subscriptions').select('token');

  if (!subs || subs.length === 0) {
    console.log('[FCM] 구독 토큰 없음');

    // Realtime fallback: push_notifications에도 저장
    await db.from('push_notifications').insert({
      title: payload.title,
      body: payload.body,
      url: payload.url || '/dashboard',
      tag: payload.tag || 'mamoru',
      read: false,
    });

    return { sent: 0, failed: 0 };
  }

  const messaging = admin.messaging(app);
  let sent = 0;
  let failed = 0;

  for (const sub of subs) {
    try {
      await messaging.send({
        token: sub.token,
        notification: {
          title: payload.title,
          body: payload.body,
        },
        webpush: {
          notification: {
            icon: '/icon-192.png',
            badge: '/icon-192.png',
            tag: payload.tag || 'mamoru',
            requireInteraction: true,
          },
          fcmOptions: {
            link: payload.url || '/dashboard',
          },
        },
        data: {
          url: payload.url || '/dashboard',
          tag: payload.tag || 'mamoru',  // SW/Realtime 중복 dedup 용
        },
      });
      sent++;
    } catch (err: unknown) {
      failed++;
      const errMsg = err instanceof Error ? err.message : String(err);
      console.error('[FCM] 발송 실패:', errMsg);
      // 토큰 만료 시 삭제
      if (errMsg.includes('not-registered') || errMsg.includes('invalid-registration')) {
        await db.from('push_subscriptions').delete().eq('token', sub.token);
      }
    }
  }

  // Realtime에도 저장 (TMS 탭 열어둔 경우 알림음)
  await db.from('push_notifications').insert({
    title: payload.title,
    body: payload.body,
    url: payload.url || '/dashboard',
    tag: payload.tag || 'mamoru',
    read: false,
  });

  console.log(`[FCM] 발송: ${sent}건 성공, ${failed}건 실패`);
  return { sent, failed };
}
