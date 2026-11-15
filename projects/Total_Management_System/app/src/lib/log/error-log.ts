import { createServiceClient } from '@/lib/supabase/server';

/**
 * 오류 모아 보기 (2026-11-15)
 *
 * 전엔 화면 · 서버 오류와 알림톡 발송 실패를 볼 곳이 Vercel 로그(개발자용)뿐이었다.
 * 새 표를 만들지 않고 이미 있는 기록 표(sync_log)에 종류를 구분해 쌓는다 — 설정 > 시스템 「오류 모아 보기」에서 본다.
 *   sync_type: error_screen(화면) · error_alimtalk(알림톡 발송 실패) · error_api(서버)
 *
 * 🔒 고객 이름 · 전화번호 같은 개인정보는 넣지 않는다(무슨 일이 실패했는지만).
 * 기록에 실패해도 원래 하던 일을 막지 않는다(절대 throw 하지 않음).
 */
export type ErrorKind = 'screen' | 'alimtalk' | 'api';

export async function logError(kind: ErrorKind, message: string): Promise<void> {
  try {
    const now = new Date().toISOString();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (createServiceClient() as any).from('sync_log').insert({
      sync_type: `error_${kind}`,
      status: 'failed',
      records_synced: 0,
      error_message: String(message || '').slice(0, 500),
      started_at: now,
      completed_at: now,
    });
  } catch {
    /* 기록 실패는 조용히 넘긴다 */
  }
}
