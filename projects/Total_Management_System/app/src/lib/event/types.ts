/**
 * EVENT(고객 접수) 타입 — 신규 테이블이라 Database 타입(never 이슈) 대신 독립 인터페이스.
 * API route에서는 (db as any) 캐스팅, 앱에서는 이 타입 사용.
 */

export type EventStatus =
  | 'received'         // 접수 (고객 제출)
  | 'payment_noticed'  // 입금안내 발송 (사장님 재고확인 후)
  | 'paid'             // 입금확인 (→ 판매 자동전환)
  | 'converted'        // 판매 전환됨 (sale_id 연결)
  | 'cancelled';

export type EventReceiveMethod = 'delivery' | 'visit'; // 택배 / 매장방문

export type EventCategoryType = 'blunt' | 'thinning' | 'long' | 'dry';

export interface EventItem {
  product_id: string | null;
  product_name: string;
  category_type: EventCategoryType;
  spec: string;            // 인치(6.0) / 감모구간 / DRY 서브타입
  slicing: boolean;        // DRY 슬라이싱 가공 여부 (+20,000)
  qty: number;
  unit_price: number;      // 슬라이싱 가공비 미포함 단가
}

export type EventCampaignType = 'stock_clearance' | 'limited' | 'group_buy' | 'tester' | 'trade_in' | 'other';

/** 수량 묶음 할인 규칙 — 같은 단가끼리, min_qty 도달 시 묶음 반복 + 나머지 정가 */
export interface DiscountRule {
  unit_price: number;   // 적용 단가 (이 단가 품목끼리 묶음)
  min_qty: number;      // 묶음 최소 수량
  bundle_price: number; // 묶음가 (min_qty 자루 묶음 가격)
}

export interface EventCampaign {
  id: string;
  name: string;
  type: EventCampaignType;
  status: 'active' | 'ended';
  is_default: boolean;
  starts_at: string | null;
  ends_at: string | null;
  memo: string | null;
  discount_rules: DiscountRule[];
  created_at: string;
  updated_at: string;
}

export const CAMPAIGN_TYPE_LABEL: Record<EventCampaignType, string> = {
  stock_clearance: '재고 전환', limited: '한정 판매', group_buy: '공동 구매',
  tester: '체험단', trade_in: '트레이드인', other: '기타',
};

export interface EventSubmission {
  id: string;
  event_number: string;
  campaign_id: string | null;
  customer_id: string | null;
  customer_name: string;
  customer_phone: string | null;
  receive_method: EventReceiveMethod;
  postcode: string | null;
  address1: string | null;
  address2: string | null;
  items: EventItem[];
  slicing_addon: number;
  total_amount: number;
  status: EventStatus;
  payment_noticed_at: string | null;
  paid_at: string | null;
  sale_id: string | null;
  memo: string | null;
  cancelled_at: string | null;
  created_at: string;
  updated_at: string;
}

export const EVENT_STATUS_LABEL: Record<EventStatus, string> = {
  received: '신규접수',
  payment_noticed: '입금대기',
  paid: '입금확인',
  converted: '판매전환',
  cancelled: '취소',
};
