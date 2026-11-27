/**
 * 복원수리 자루 수 세기 (2026-11-27)
 *
 * 판매 · 납품의 복원수리 줄(category='RS')은 금액은 전부 복원수리 매출이지만, 자루 수는 「가위」 줄만 센다.
 * 전엔 배송비만 빼고 다 셌다 → 날 변형 교정 같은 추가 금액 줄이 1자루씩 더해져 대시보드 자루 수가 부풀었다.
 *
 * 자루가 아닌 줄
 *  - 배송비 · 배송 · 각인 (운영 자료에 실제로 있는 이름)
 *  - 「추가작업 · …」 — 판매 입력 「추가 작업」 · 거래처 매출 「추가 항목」이 이 이름으로 저장한다
 * 예전에 자유롭게 적은 이름(가위수리 · 수리 · 복원수리(날변형) 등)은 대부분 실제 가위라 그대로 센다.
 */
export const REPAIR_EXTRA_PREFIX = '추가작업 · ';

const NON_SCISSOR_NAMES = new Set(['배송비', '배송', '각인']);

/** 이 복원수리 줄이 가위가 아닌(금액만 더하는) 줄인가 */
export function isNonScissorRepairLine(productName: string | null | undefined): boolean {
  const n = (productName || '').trim();
  return NON_SCISSOR_NAMES.has(n) || n.startsWith(REPAIR_EXTRA_PREFIX.trim());
}

/** 자루 수 — 가위가 아닌 줄은 0, 나머지는 받은 수량 그대로 */
export function scissorQty(productName: string | null | undefined, qty: number): number {
  return isNonScissorRepairLine(productName) ? 0 : qty;
}

/** 추가 작업 줄의 저장 이름 (이미 붙어 있거나 배송비면 그대로) */
export function repairExtraName(name: string): string {
  const n = name.trim();
  if (NON_SCISSOR_NAMES.has(n) || n.startsWith(REPAIR_EXTRA_PREFIX.trim())) return n;
  return REPAIR_EXTRA_PREFIX + n;
}
