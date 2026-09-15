/**
 * 택배사 단일 출처 (150, 2026-09-15)
 *
 * 문제: 시스템 전체가 "송장번호 = 롯데 ALPS 송장" 이라는 암묵적 전제 위에 있었다.
 *   롯데 집하취소하고 우체국으로 직접 보낸 건을 기록할 방법이 없어서
 *   ① 죽은 롯데 송장이 그대로 남거나 ② 우체국 번호를 롯데 송장으로 오인해
 *   크론이 매시간 헛조회하고 배송완료가 영영 안 떴다.
 *
 * 이제 "롯데 ALPS 접수"와 "실제로 어떻게 보냈는가"를 분리한다.
 * 자동추적(집하감지·배송완료) 대상 판정은 **반드시 이 파일의 isAlpsTrackable 만** 쓴다.
 */

/** 직접 전달(택배 없이 수령·방문 전달) — 송장번호가 없는 출고 */
export const COURIER_DIRECT = '직접전달';
/** 롯데 ALPS 연동 택배사 — 송장 생성·자동추적이 되는 유일한 값 */
export const COURIER_LOTTE = '롯데택배';

/** 수동 입력 시 고를 수 있는 택배사. 첫 항목이 기본값 */
export const COURIER_OPTIONS = [
  COURIER_LOTTE,
  '우체국택배',
  'CJ대한통운',
  '한진택배',
  '로젠택배',
  '경동택배',
  COURIER_DIRECT,
] as const;

export type CourierName = (typeof COURIER_OPTIONS)[number];

/**
 * ALPS 자동추적(집하감지 → 출고완료 / 인수자등록 → 배송완료) 대상인가?
 *
 * 🔴 NULL·빈값은 **true**. 구데이터 호환 — 이 컬럼이 생기기 전 송장은 전부 롯데였다.
 *    여기서 false 를 돌리면 기존 건들이 통째로 자동추적에서 빠진다.
 */
export function isAlpsTrackable(courierName?: string | null): boolean {
  if (!courierName) return true;
  return courierName.trim() === COURIER_LOTTE;
}

/** 택배 없이 직접 전달한 건인가 (송장번호 없음) */
export function isDirectHandover(courierName?: string | null): boolean {
  return (courierName || '').trim() === COURIER_DIRECT;
}

/** 화면 표기용 — 값이 없으면 롯데택배(구데이터) */
export function courierLabel(courierName?: string | null): string {
  return (courierName || '').trim() || COURIER_LOTTE;
}
