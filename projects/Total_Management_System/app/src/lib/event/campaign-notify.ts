/**
 * EVENT 알림톡 공용 변수 — 이벤트명(event_name) · 고객 안내 문구(event_notice)
 *
 * 범용 템플릿(EVENT_신청완료 / 입금확인 / 출고완료)이 특정 이벤트에 묶이지 않도록 캠페인에서 읽어 채운다.
 * - event_name   : event_campaigns.name
 * - event_notice : event_campaigns.customer_notice (145) — 이벤트별 진행 안내(주소 회신·발송 기한 등). 신청완료에서만 사용
 *
 * 🔴 알림톡 변수가 비면 솔라피가 문자로 대체 발송하므로 두 값 모두 항상 기본값으로 채운다.
 *    마이그 145 실행 전(컬럼 없음)에도 깨지지 않게 select('*') 로 읽는다.
 */

export const DEFAULT_EVENT_NAME = 'MAMORU 이벤트';
export const DEFAULT_EVENT_NOTICE = '추가로 필요한 사항이 있으면 따로 연락드립니다';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export async function getEventCampaignNotifyVars(
  db: Db,
  campaignId: string | null | undefined,
): Promise<{ event_name: string; event_notice: string }> {
  let name = '';
  let notice = '';
  if (campaignId) {
    const { data } = await db.from('event_campaigns').select('*').eq('id', campaignId).maybeSingle();
    name = typeof data?.name === 'string' ? data.name.trim() : '';
    // 템플릿에서 "• #{event_notice}" 한 줄로 쓰이므로 줄바꿈은 공백으로 합친다
    notice = typeof data?.customer_notice === 'string' ? data.customer_notice.replace(/\s*\n\s*/g, ' ').trim() : '';
  }
  return {
    event_name: name || DEFAULT_EVENT_NAME,
    event_notice: notice || DEFAULT_EVENT_NOTICE,
  };
}

/** 마이그 145(customer_notice) 미실행 시 PostgREST '컬럼 없음' 에러 판별 — 캠페인 저장 API 가 안내 메시지로 바꿔 응답 */
export function isMissingNoticeColumn(err: { code?: string; message?: string } | null | undefined): boolean {
  return !!err && (err.code === 'PGRST204' || /customer_notice/.test(err.message || ''));
}

/** 판매(offline_sales) → 그 판매로 전환된 EVENT 접수의 캠페인 변수 (출고 알림용) */
export async function getEventNotifyVarsBySale(db: Db, saleId: string): Promise<{ event_name: string; event_notice: string }> {
  const { data } = await db.from('event_submissions').select('campaign_id').eq('sale_id', saleId).limit(1).maybeSingle();
  return getEventCampaignNotifyVars(db, data?.campaign_id);
}
