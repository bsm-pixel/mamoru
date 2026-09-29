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
  /** 휴대폰이 2분 안에 못 받으면 사장님 메일 예비발송 (기본 true) */
  fallback?: boolean;
}

/**
 * 등록된 모든 디바이스에 FCM 푸시 발송 (무조건 발송 — on/off 게이팅 없음).
 * 🔴 고객 접수/행동 알림은 사장님이 놓치면 안 되므로 어떤 설정으로도 차단하지 않는다.
 *    (2026-08-01: isPushEnabled 게이팅 제거 — 접수 알림 오락가락 근본원인 정리)
 * 발송은 sendEach 배치 1회로 처리 → 순차 send 루프의 지연/부분누락 제거.
 *
 * 🔴 2026-09-29 수신 확인 + 메일 예비발송 (마이그 155) — "로그인 풀린 사이 알림 누락" 신고
 *    ① 기록을 먼저 만들고 그 id(nid)를 푸시에 실어 보낸다 → 휴대폰 SW가 받는 즉시 /api/push/ack
 *    ② 2분 안에 휴대폰 수신이 없으면 크론(push-fallback)이 사장님 메일 1통(bsm@mamoru.kr → Gmail 앱 알림)
 *    ③ 휴대폰이 0대 / 휴대폰 발송 전부 실패 / 크론 중단 → 기다리지 않고 즉시 메일
 *    ④ Urgency: high — 안드로이드 절전(Doze) 중에도 지연 없이 깨워서 표시
 */
export async function sendPushToAll(payload: PushPayload): Promise<{ sent: number; failed: number }> {
  const { createServiceClient } = await import('@/lib/supabase/server');
  const { isMobileLabel, isSweepAlive, claimAndSendFallback, sendFallbackUnclaimed } = await import('./push-fallback');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;

  const url = payload.url || '/dashboard';
  const tag = payload.tag || 'mamoru';
  const wantFallback = payload.fallback !== false;

  const { data: subsData } = await db.from('push_subscriptions').select('token, device_info');
  const subs: Array<{ token: string; device_info: string | null }> = subsData || [];
  const mobileCount = subs.filter((s) => isMobileLabel(s.device_info)).length;

  // ① 기록 먼저 (Realtime 폴백 겸용 — TMS 탭 열려있으면 소리). 마이그 155 전이면 기본 컬럼만.
  const base = { title: payload.title, body: payload.body, url, tag, read: false };
  let nid: string | null = null;
  {
    const full = await db.from('push_notifications')
      .insert({ ...base, target_count: subs.length, mobile_target_count: mobileCount, fallback_needed: wantFallback })
      .select('id').single();
    if (!full.error) nid = full.data?.id ?? null;
    else {
      const legacy = await db.from('push_notifications').insert(base).select('id').single();
      nid = legacy.data?.id ?? null;
    }
  }

  const app = getApp();
  let sent = 0;
  let failed = 0;
  let mobileSent = 0;

  if (app && subs.length > 0) {
    const messaging = admin.messaging(app);
    const messages = subs.map((s) => ({
      token: s.token,
      notification: { title: payload.title, body: payload.body },
      webpush: {
        // 절전 중인 안드로이드도 즉시 깨워 표시. 없으면 normal 로 취급돼 몇 분~몇 시간 묶일 수 있다
        headers: { Urgency: 'high' },
        notification: {
          icon: '/icon-192.png',
          badge: '/icon-192.png',
          tag,
          requireInteraction: true,
        },
        fcmOptions: { link: url },
      },
      data: {
        url,
        tag, // SW/Realtime 중복 dedup 용
        nid: nid || '', // SW 수신 확인(ack) 용
      },
    }));

    try {
      const resp = await messaging.sendEach(messages);
      sent = resp.successCount;
      failed = resp.failureCount;
      // 만료/무효 토큰만 정리
      const stale: string[] = [];
      resp.responses.forEach((r, i) => {
        if (r.success) {
          if (isMobileLabel(subs[i].device_info)) mobileSent++;
          return;
        }
        const code = r.error?.code || '';
        console.error('[FCM] 발송 실패:', code, r.error?.message);
        if (code.includes('not-registered') || code.includes('invalid-registration') || code.includes('invalid-argument')) {
          stale.push(subs[i].token);
        }
      });
      if (stale.length) await db.from('push_subscriptions').delete().in('token', stale);
    } catch (err) {
      failed = messages.length;
      console.error('[FCM] sendEach 오류:', err instanceof Error ? err.message : String(err));
    }
  } else if (!app) {
    console.warn('[FCM] Firebase Admin 미설정 — 푸시 생략, 메일 예비발송만 판단');
  } else {
    console.log('[FCM] 구독 토큰 없음');
  }

  if (nid) {
    await db.from('push_notifications').update({ sent_count: sent, failed_count: failed }).eq('id', nid)
      .then(() => {}, () => {});
  }

  // ③ 기다려도 휴대폰에 안 올 게 확실하거나, 2분 뒤 확인해줄 크론이 죽어 있으면 → 즉시 메일
  if (wantFallback) {
    const reason =
      mobileCount === 0 ? '휴대폰 알림 미등록 — TMS 앱을 열어 알림을 허용해주세요'
      : mobileSent === 0 ? '휴대폰 발송 실패 — TMS 앱을 한 번 열어주세요'
      : !(await isSweepAlive(db)) ? '수신확인 점검 중단 — 즉시 발송'
      : null;
    if (reason) {
      const row = { id: nid || '', title: payload.title, body: payload.body, url };
      const r = nid ? await claimAndSendFallback(db, row, reason) : 'no-column';
      if (r === 'no-column') await sendFallbackUnclaimed(row, reason);
    }
  }

  console.log(`[FCM] 발송: ${sent}건 성공(휴대폰 ${mobileSent}), ${failed}건 실패`);
  return { sent, failed };
}
