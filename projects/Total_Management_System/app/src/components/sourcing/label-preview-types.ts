/**
 * § 1688 사입 — 디자인 모니터 데모 전용 타입
 *
 * Phase A (디자인 모니터) → Phase B (운영) 이전 시 동일 인터페이스로 사용 가능하도록
 * 운영 DB 스키마(`purchase_orders` / `purchase_order_items`)와 키 이름을 맞춤.
 *
 * ※ Phase A 단계에서는 useState 로컬 메모리만 사용. API/Supabase 호출 없음.
 */

export type InspectionStatus = 'pending' | 'matched' | 'promoted' | 'rejected';

export type DemoPOItem = {
  id: string;
  vendor_url: string;
  product_name: string;
  features_memo: string;
  moq: number | null;
  unit_price: number; // CNY
  quantity: number;
  sticker_no: string; // PO-2026XXXX-001 형식
  inbound_photos: string[]; // base64 dataURL 또는 mock URL
  inbound_memo: string;
  inspection_status: InspectionStatus;
  promoted_sku?: string;
  promoted_name?: string;
};

export type DemoPO = {
  po_number: string;
  supplier_name: string;
  supplier_url: string;
  order_date: string; // YYYY-MM-DD
  exchange_rate: number;
  items: DemoPOItem[];
};

export const STATUS_LABEL: Record<InspectionStatus, string> = {
  pending: '대기',
  matched: '매칭완료',
  promoted: '정식등록',
  rejected: '보류',
};

export const STATUS_TONE: Record<InspectionStatus, string> = {
  pending: 'bg-stone-100 text-stone-500 border-stone-200',
  matched: 'bg-emerald-50 text-emerald-700 border-emerald-200',
  promoted: 'bg-stone-900 text-white border-stone-900',
  rejected: 'bg-rose-50 text-rose-600 border-rose-200',
};
