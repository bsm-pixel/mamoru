import { describe, it, expect } from 'vitest';
import { saleNet, saleOutstanding, deliveryNet, deliveryOutstanding } from '@/lib/sales/amounts';
import { isValidRepairTransition, REPAIR_STATUS_ORDER, getAllowedRepairTransitions } from '@/lib/repair/transitions';

/** 계산·상태 규칙 — 화면을 고치다 공식이 바뀌면 여기서 잡힌다 */
describe('금액 공식', () => {
  it('판매: 순액 = 합계 − 할인, 미수 = 순액 − 받은 돈 (0 아래로 안 내려감)', () => {
    const s = { total_amount: 300000, discount_amount: 20000, paid_amount: 100000 };
    expect(saleNet(s)).toBe(280000);
    expect(saleOutstanding(s)).toBe(180000);
    expect(saleOutstanding({ total_amount: 100, discount_amount: 0, paid_amount: 500 })).toBeGreaterThanOrEqual(0);
  });
  it('납품: 합계가 이미 할인 반영값이라 할인을 다시 빼지 않는다', () => {
    const d = { total_amount: 500000, discount_amount: 50000, paid_amount: 200000 };
    expect(deliveryNet(d)).toBe(500000);
    expect(deliveryOutstanding(d)).toBe(300000);
  });
});

describe('복원수리 상태 규칙', () => {
  it('끝난 상태(완료·취소)에서는 어디로도 못 간다', () => {
    expect(getAllowedRepairTransitions('completed' as never)).toEqual([]);
    expect(getAllowedRepairTransitions('cancelled' as never)).toEqual([]);
  });
  it('진행 순서의 모든 단계는 다음 단계로 가는 길이 하나는 있다', () => {
    for (let i = 0; i < REPAIR_STATUS_ORDER.length - 1; i++) {
      const from = REPAIR_STATUS_ORDER[i];
      const later = REPAIR_STATUS_ORDER.slice(i + 1);
      expect(later.some((to) => isValidRepairTransition(from, to)), `${from} 에서 앞으로 가는 길이 없음`).toBe(true);
    }
  });
  it('뒤로 되돌리는 변경은 없다', () => {
    for (let i = 1; i < REPAIR_STATUS_ORDER.length; i++) {
      for (let j = 0; j < i; j++) {
        expect(isValidRepairTransition(REPAIR_STATUS_ORDER[i], REPAIR_STATUS_ORDER[j]), `${REPAIR_STATUS_ORDER[i]} → ${REPAIR_STATUS_ORDER[j]}`).toBe(false);
      }
    }
  });
});
