/**
 * 푸시 "휴대폰 미수신 → 사장님 메일 예비발송" (2026-09-29, 마이그 155)
 * (컬럼명 fallback_* 는 채널 중립 — 처음 문자로 설계했다가 메일로 확정)
 *
 * 흐름:
 *   sendPushToAll → push_notifications 행(fallback_needed=true) + FCM 발송(data.nid)
 *   휴대폰 서비스워커가 받는 순간 → /api/push/ack → acked_at 기록
 *   크론(1분) → 2분 지나도 acked_at 없음 → 사장님 메일 1통 (fallback_sent_at 선점으로 1회만)
 *
 * 즉시 메일(2분 안 기다림):
 *   - 휴대폰이 한 대도 등록돼 있지 않음 / 휴대폰 발송이 전부 실패 → 기다려도 안 온다
 *   - 크론 심장박동이 끊김 → 2분 뒤 확인해줄 주체가 없다 (안전망의 안전망)
 */

import { sendOwnerAlert } from '@/lib/notification/owner-alert';

/** 휴대폰 수신 확인을 기다리는 시간. 넘으면 메일 */
export const ACK_GRACE_MS = 2 * 60 * 1000;
/** 이 시간보다 오래된 미수신 건은 크론이 더 이상 메일로 보내지 않는다(크론 장기 중단 후 폭주 방지) */
const SWEEP_WINDOW_MS = 60 * 60 * 1000;
/** 크론 심장박동 — system_settings 키 */
export const HEARTBEAT_KEY = 'push.fallback_heartbeat';
const HEARTBEAT_STALE_MS = 5 * 60 * 1000;

/** "휴대폰"으로 인정하는 기기 — 수신 확인은 휴대폰에서 온 것만 인정한다 (PC만 받고 폰이 못 받으면 메일) */
export function isMobileLabel(label: string | null | undefined): boolean {
  return /Android|iPhone|iPad/i.test(label || '');
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

interface PendingRow {
  id: string;
  title: string;
  body: string;
  url?: string | null;
}

/** Gmail 알림에 제목이 먼저 보이므로 핵심(무엇·누구)을 제목에 다 넣는다 */
function sendAlert(row: PendingRow, reason: string) {
  return sendOwnerAlert(
    `[MAMORU] ${row.title} · ${row.body}`,
    `${row.title}\n${row.body}\n\n앱 알림이 휴대폰에 도착하지 않아 메일로 보냅니다.\n(${reason})`,
    row.url || undefined,
  );
}

/**
 * 메일 1통 — fallback_sent_at 을 먼저 선점해 크론·즉시경로가 동시에 와도 1회만 발송.
 * 반환: 'sent' | 'failed' | 'skipped'(이미 수신/발송됨) | 'no-column'(마이그 155 미적용)
 */
export async function claimAndSendFallback(
  db: Db,
  row: PendingRow,
  reason: string,
): Promise<'sent' | 'failed' | 'skipped' | 'no-column'> {
  const now = new Date().toISOString();
  const { data: claimed, error } = await db
    .from('push_notifications')
    .update({ fallback_sent_at: now })
    .eq('id', row.id)
    .is('fallback_sent_at', null)
    .is('acked_at', null)
    .select('id');

  if (error) return 'no-column';
  if (!claimed || claimed.length === 0) return 'skipped';

  const r = await sendAlert(row, reason);
  await db.from('push_notifications').update({ fallback_result: r.result }).eq('id', row.id);
  return r.ok ? 'sent' : 'failed';
}

/** 마이그 155 전이라 선점을 못 할 때 — 즉시경로에서만 쓰는 무조건 발송 */
export async function sendFallbackUnclaimed(row: PendingRow, reason: string) {
  return sendAlert(row, reason);
}

/** 크론 심장박동이 살아 있는가 — 죽었으면 2분 뒤 확인이 안 되므로 즉시 메일로 간다 */
export async function isSweepAlive(db: Db): Promise<boolean> {
  try {
    const { data } = await db.from('system_settings').select('value').eq('key', HEARTBEAT_KEY).maybeSingle();
    const at = data?.value ? new Date(String(data.value).replace(/^"|"$/g, '')).getTime() : 0;
    return !!at && Date.now() - at < HEARTBEAT_STALE_MS;
  } catch {
    return false;
  }
}

/** 크론 1회분: 심장박동 기록 → 2분 지난 미수신 건 메일 */
export async function runFallbackSweep(db: Db): Promise<{ checked: number; sent: number; failed: number; error?: string }> {
  const nowIso = new Date().toISOString();
  await db.from('system_settings').upsert(
    { key: HEARTBEAT_KEY, value: nowIso, updated_at: nowIso },
    { onConflict: 'key' },
  );

  const due = new Date(Date.now() - ACK_GRACE_MS).toISOString();
  const oldest = new Date(Date.now() - SWEEP_WINDOW_MS).toISOString();
  const { data: rows, error } = await db
    .from('push_notifications')
    .select('id, title, body, url')
    .eq('fallback_needed', true)
    .is('acked_at', null)
    .is('fallback_sent_at', null)
    .lte('created_at', due)
    .gte('created_at', oldest)
    .order('created_at', { ascending: true })
    .limit(20);

  if (error) return { checked: 0, sent: 0, failed: 0, error: error.message };

  let sent = 0;
  let failed = 0;
  for (const row of (rows || []) as PendingRow[]) {
    const r = await claimAndSendFallback(db, row, '2분 내 휴대폰 수신 확인 없음');
    if (r === 'sent') sent++;
    else if (r === 'failed') failed++;
  }
  return { checked: rows?.length || 0, sent, failed };
}
