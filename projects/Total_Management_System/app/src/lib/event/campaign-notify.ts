/**
 * EVENT 알림톡 공용 변수 — 이벤트명 · 신청 항목 · 진행 안내
 *
 * 범용 템플릿(EVENT_신청완료 / 입금확인 / 출고완료)이 특정 이벤트에 묶이지 않도록 캠페인에서 읽어 채운다.
 *
 * 🔴 2026-09-16 확장 — 무료(증정·체험단) 이벤트 지원 (마이그 152)
 *    전엔 본문에 `결제 금액`·`받으실 곳`·`입금 계좌`가 **고정 텍스트**로 박혀 있어
 *    증정 이벤트인데 "결제 금액 0원"이 떠서 고객이 "왜 결제가?" 하게 됐다(사장님 지적).
 *
 *    알림톡은 **변수 안의 변수를 치환하지 않는다**(치환은 1회). 그래서 주소·금액·계좌를 끼워 넣으려면
 *    서버가 문자열을 **완성해서** 한 변수로 넘겨야 한다 → 그게 event_detail 이다.
 *
 *    ⚠️ 본문을 통째로 변수로 만들면 카카오 심사에서 "치환문구로만 구성"으로 반려될 수 있다.
 *       그래서 인사·신청내역·안내 골격은 본문에 남기고 **가운데 한 블록만** 변수로 뺀다.
 *
 * 변수
 * - event_name   : event_campaigns.name
 * - items        : 캠페인 items_label 이 있으면 그 문구, 없으면 "제품명 N개" 줄바꿈 목록
 * - event_detail : 결제·수령·계좌·진행안내를 상황에 맞게 조립한 한 덩어리 (신청완료에서만 사용)
 * - event_notice : (구) 캠페인 안내 문구 한 줄 — 옛 템플릿 호환용으로 계속 내보낸다
 *
 * 🔴 알림톡 변수가 비면 솔라피가 문자로 대체 발송(버튼도 사라짐)하므로 모든 값을 항상 채운다.
 *    마이그 152 실행 전(컬럼 없음)에도 깨지지 않게 select('*') 로 읽고 기본값으로 떨어진다.
 */

export const DEFAULT_EVENT_NAME = 'MAMORU 이벤트';
export const DEFAULT_EVENT_NOTICE = '추가로 필요한 사항이 있으면 따로 연락드립니다';

/**
 * 입금 계좌 — 전엔 알림톡 본문에 고정 텍스트로 있어서 계좌가 바뀌면 **본문 수정 + 재검수**가 필요했다.
 * 변수 안으로 옮겨서, 계좌 변경은 이 상수만 고치면 되고 무료 이벤트엔 아예 안 나간다.
 */
export const EVENT_BANK_ACCOUNTS = [
  '우리 1002-439-462514 백성민',
  '신한 110-445-097604 마모루(백성민)',
];

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface EventCampaignConfig {
  event_name: string;
  event_notice: string;
  /** 'paid' = 입금 필요 / 'free' = 증정·체험단 */
  payment_type: 'paid' | 'free';
  /** #{items} 대체 문구 (비면 제품명 사용) */
  items_label: string;
}

/** 캠페인 설정 읽기 — 마이그 152 미적용/캠페인 없음이면 전부 기존 동작(유료)으로 떨어진다 */
export async function getEventCampaignConfig(db: Db, campaignId: string | null | undefined): Promise<EventCampaignConfig> {
  let name = '';
  let notice = '';
  let paymentType: 'paid' | 'free' = 'paid';
  let itemsLabel = '';

  if (campaignId) {
    const { data } = await db.from('event_campaigns').select('*').eq('id', campaignId).maybeSingle();
    name = typeof data?.name === 'string' ? data.name.trim() : '';
    // 옛 템플릿에서 "• #{event_notice}" 한 줄로 쓰이므로 줄바꿈은 공백으로 합친다
    notice = typeof data?.customer_notice === 'string' ? data.customer_notice.replace(/\s*\n\s*/g, ' ').trim() : '';
    if (data?.payment_type === 'free') paymentType = 'free';
    itemsLabel = typeof data?.items_label === 'string' ? data.items_label.trim() : '';
  }

  return {
    event_name: name || DEFAULT_EVENT_NAME,
    event_notice: notice || DEFAULT_EVENT_NOTICE,
    payment_type: paymentType,
    items_label: itemsLabel,
  };
}

/** 이전 이름 유지 — 호출부(입금안내·입금확인)가 이름·안내문구만 쓸 때 */
export async function getEventCampaignNotifyVars(
  db: Db,
  campaignId: string | null | undefined,
): Promise<{ event_name: string; event_notice: string }> {
  const c = await getEventCampaignConfig(db, campaignId);
  return { event_name: c.event_name, event_notice: c.event_notice };
}

export interface EventItemLike {
  product_name: string;
  qty: number;
  slicing?: boolean;
}

/** #{items} — 캠페인이 표기를 지정했으면 그 문구, 아니면 제품명+수량 (내부 품목 데이터는 그대로) */
export function buildItemsText(items: EventItemLike[], itemsLabel: string): string {
  if (itemsLabel) return itemsLabel;
  const text = items
    .map((it) => `${it.product_name}${it.slicing ? '(슬라이싱)' : ''} ${it.qty}개`)
    .join('\n');
  return text || '신청 내역 확인 중';
}

export interface EventDetailInput {
  totalAmount: number;
  /** 매장 방문 수령 여부 */
  isVisit: boolean;
  /** 배송지 (isVisit 이면 무시) */
  address: string;
}

/**
 * #{event_detail} — 결제·수령·계좌·진행안내를 상황에 맞게 조립한다.
 *
 *  유료 : 금액 + 수령지 + 계좌 + "입금이 확인되면 바로 준비를 시작합니다"
 *  무료 : 수령지만 + 캠페인 안내문구(없으면 "바로 준비를 시작합니다")
 *         → 금액·계좌 줄이 **아예 생기지 않는다**. 증정인데 0원이 뜨던 문제의 해결점
 */
export function buildEventDetail(cfg: EventCampaignConfig, input: EventDetailInput): string {
  const lines: string[] = [];
  const isPaid = cfg.payment_type === 'paid';

  // 금액 — 유료이고 실제 금액이 있을 때만. 0원짜리 유료는 금액 줄을 만들지 않는다
  if (isPaid && input.totalAmount > 0) {
    lines.push(`• 결제 금액 : ${input.totalAmount.toLocaleString('ko-KR')}원`);
  }

  // 수령 방법 / 배송지
  lines.push(input.isVisit ? '• 수령 방법 : 매장 방문 수령' : `• 받으실 곳 : ${input.address || '접수하신 주소'}`);

  // 계좌 — 유료 + 실제 받을 금액이 있을 때만
  if (isPaid && input.totalAmount > 0) {
    lines.push('', '입금 계좌', ...EVENT_BANK_ACCOUNTS.map((a) => `• ${a}`));
  }

  // 마무리 안내
  lines.push('');
  if (isPaid && input.totalAmount > 0) {
    lines.push('입금이 확인되면 바로 준비를 시작합니다');
    if (cfg.event_notice && cfg.event_notice !== DEFAULT_EVENT_NOTICE) lines.push(cfg.event_notice);
  } else {
    // 무료: 캠페인이 안내문구를 지정했으면 그것만(예: "선정 결과는 마감 후 개별 안내드립니다")
    lines.push(cfg.event_notice && cfg.event_notice !== DEFAULT_EVENT_NOTICE
      ? cfg.event_notice
      : '바로 준비를 시작합니다');
  }

  return lines.join('\n').trim();
}

/** 마이그 145(customer_notice)·152(payment_type/items_label) 미실행 시 '컬럼 없음' 판별 — 캠페인 저장 API 가 안내로 바꿔 응답 */
export function isMissingNoticeColumn(err: { code?: string; message?: string } | null | undefined): boolean {
  return !!err && (err.code === 'PGRST204' || /customer_notice|payment_type|items_label/.test(err.message || ''));
}

/** 판매(offline_sales) → 그 판매로 전환된 EVENT 접수의 캠페인 변수 (출고 알림용) */
export async function getEventNotifyVarsBySale(db: Db, saleId: string): Promise<{ event_name: string; event_notice: string }> {
  const { data } = await db.from('event_submissions').select('campaign_id').eq('sale_id', saleId).limit(1).maybeSingle();
  return getEventCampaignNotifyVars(db, data?.campaign_id);
}

/** 판매 → 그 EVENT 접수 캠페인이 무료인지 (입금확인 알림톡 생략 판단용) */
export async function getEventCampaignBySubmission(db: Db, campaignId: string | null | undefined) {
  return getEventCampaignConfig(db, campaignId);
}
