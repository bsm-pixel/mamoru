/**
 * 배송상태 파생 — 판매(offline_sales) / 납품(deliveries) 공용 (2026-07-16)
 *
 * "상태"(결제/전체)와 별개로, 출고 흐름만 떼어낸 라벨.
 *   판매(B2C): 배송대기 → 배송중 → 배송완료 / 직접수령  (2026-08-23 주문관리와 워딩 통일)
 *   납품(B2B): 출고대기 → 출고완료 (도매 특성상 '출고' 용어 유지)
 *
 * ⚠️ 판정 기준은 상세 패널·getRowStateSale 과 100% 동일 필드(delivered_at/shipped_at/invoice_number)를
 *    재사용한다. 새 기준을 만들지 않는다(라벨 드리프트 0).
 */
export type ShipStatus = {
  label: string;
  /** amber = 사장님이 아직 할 일 있음(출고대기) / green = 진행·완료 / mute = 해당 없음 */
  tone: 'amber' | 'green' | 'mute';
  /** 롯데 기사님 수거 자동감지로 출고 처리된 건 (보조표기) */
  autoPicked?: boolean;
};

const NONE: ShipStatus = { label: '—', tone: 'mute' };

/** 판매(offline_sales) 배송상태 */
export function getSaleShipStatus(sale: {
  invoice_number?: string | null;
  shipped_at?: string | null;
  delivered_at?: string | null;
  shipped_source?: string | null;
}): ShipStatus {
  if (sale.delivered_at && !sale.invoice_number) return { label: '직접수령', tone: 'green' };
  if (sale.delivered_at) return { label: '배송완료', tone: 'green' };
  if (sale.shipped_at) return { label: '배송중', tone: 'green', autoPicked: sale.shipped_source === 'alps_pickup' };
  if (sale.invoice_number) return { label: '배송대기', tone: 'amber' };
  return NONE;
}

/** 납품(deliveries) 배송상태 — status 기반 (getRowStateDelivery 와 같은 기준) */
export function getDeliveryShipStatus(dl: {
  status?: string | null;
  tracking_number?: string | null;
  delivered_at?: string | null;
  cancelled_at?: string | null;
}): ShipStatus {
  if (dl.cancelled_at) return NONE;
  if (dl.delivered_at) return { label: '배송완료', tone: 'green' };
  if (dl.status === 'shipped' || dl.status === 'settled') return { label: '출고완료', tone: 'green' };
  if (dl.status === 'confirmed' && dl.tracking_number) return { label: '출고대기', tone: 'amber' };
  return NONE;
}
