import { describe, it, expect } from 'vitest';
import { canHandoverWithSale } from '@/lib/repair/transitions';

// 2026-11-22 판매 [고객 수령 완료] 때 함께 맡은 가위를 같이 닫아도 되는 조건 — 서버 · 화면이 같이 쓴다
describe('함께 맡은 가위를 판매와 같이 건넬 수 있는가', () => {
  it('작업완료 · 작업중이고 송장이 없으면 같이 건넨다', () => {
    expect(canHandoverWithSale({ status: 'ready_to_ship', invoice_number: null, proceed_type: '출장수거' })).toBe(true);
    expect(canHandoverWithSale({ status: 'repairing', invoice_number: null, proceed_type: '출장수거' })).toBe(true);
  });
  it('아직 작업 전이면 남긴다', () => {
    expect(canHandoverWithSale({ status: 'intake', invoice_number: null })).toBe(false);
    expect(canHandoverWithSale({ status: 'cost_notified', invoice_number: null })).toBe(false);
  });
  it('송장이 있으면 택배로 나갈 건이라 건드리지 않는다', () => {
    expect(canHandoverWithSale({ status: 'ready_to_ship', invoice_number: '123456789012' })).toBe(false);
  });
  it('이미 끝났거나 취소된 건은 건드리지 않는다', () => {
    for (const status of ['shipped', 'delivered', 'completed', 'cancelled']) {
      expect(canHandoverWithSale({ status, invoice_number: null })).toBe(false);
    }
  });
  it('매장 방문 건은 「고객 전달 완료」로 따로 끝난다', () => {
    expect(canHandoverWithSale({ status: 'ready_to_ship', invoice_number: null, proceed_type: '직접방문' })).toBe(false);
  });
});
