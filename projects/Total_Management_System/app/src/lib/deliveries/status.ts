/**
 * B2B 납품 상태 표시 — 단일 출처 (110, 2026-07-12)
 *
 * 문제: 뱃지가 `status` 하나로만 결정돼서
 *   · 송장만 발급했는데 "출고완료" 로 떴고 (api/lotte/book 이 status='shipped' 를 강제했음)
 *   · 배송이 끝나도(delivered_at) 계속 "출고완료" 로 남았다
 *
 * 이제 4단계로 정확히 나눈다. B2C 판매(offline_sales)와 같은 정의:
 *   납품확정 → (송장발급) 출고대기 → (기사님 수거=집하) 출고완료 → (인수자등록) 배송완료
 *
 * ⚠️ 상세 패널·목록·판매 통합목록이 전부 이 함수를 써야 한다. 라벨 규칙이 갈라지면 안 됨.
 */
export interface DeliveryStatusInput {
  status?: string | null;
  tracking_number?: string | null;
  delivered_at?: string | null;
  shipped_source?: string | null;
  cancelled_at?: string | null;
}

export interface DeliveryStatusChip {
  label: string;
  className: string;
  /** 롯데 기사님 수거 스캔으로 자동 처리된 건 (보조문구용) */
  autoPicked: boolean;
}

export function getDeliveryStatusChip(dl: DeliveryStatusInput): DeliveryStatusChip {
  const autoPicked = dl.shipped_source === 'alps_pickup';

  if (dl.cancelled_at) {
    return { label: '취소', className: 'bg-red-100 text-red-700', autoPicked: false };
  }
  // 배송완료 (ALPS 인수자등록 자동 감지) — 지금까지 뱃지에 아예 안 뜨던 단계
  if (dl.delivered_at) {
    return { label: '배송완료', className: 'bg-neutral-100 text-neutral-600', autoPicked };
  }
  // 출고완료 = 기사님이 실제로 수거해 감 (집하)
  if (dl.status === 'shipped' || dl.status === 'settled') {
    return { label: '출고완료', className: 'bg-green-100 text-green-700', autoPicked };
  }
  // 출고대기 = 송장은 발급됐지만 기사님이 아직 안 옴  ← 예전엔 이게 "출고완료" 로 잘못 떴다
  if (dl.status === 'confirmed' && dl.tracking_number) {
    return { label: '출고대기', className: 'bg-amber-100 text-amber-700', autoPicked: false };
  }
  if (dl.status === 'confirmed') {
    return { label: '납품확정', className: 'bg-blue-100 text-blue-700', autoPicked: false };
  }
  return { label: '작성중', className: 'bg-neutral-100 text-neutral-600', autoPicked: false };
}

/** 송장은 있는데 아직 수거 전 → "기사님 수거하면 자동 출고완료" 안내를 띄울 상태 */
export function isAwaitingPickup(dl: DeliveryStatusInput): boolean {
  return !dl.cancelled_at && dl.status === 'confirmed' && !!dl.tracking_number;
}
