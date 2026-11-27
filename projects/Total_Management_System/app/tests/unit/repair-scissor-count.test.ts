import { describe, it, expect } from 'vitest';
import { scissorQty, isNonScissorRepairLine, repairExtraName } from '@/lib/repair/scissor-count';

/**
 * 복원수리 자루 수 (2026-11-27) — 금액만 더하는 줄(배송비 · 각인 · 추가작업)은 자루로 세지 않는다.
 * 전엔 배송비만 빼서 「날 변형 교정」 같은 추가 금액 줄이 1자루씩 더해졌다(대시보드 자루 수 부풀림).
 */
describe('복원수리 자루 수', () => {
  it('가위 줄은 수량 그대로', () => {
    expect(scissorQty('복원수리 (마모루)', 3)).toBe(3);
    expect(scissorQty('복원수리 (타사제품)', 2)).toBe(2);
    expect(scissorQty('복원수리', 12)).toBe(12);
    // 예전에 자유롭게 적은 이름도 실제 가위라 그대로 센다
    expect(scissorQty('가위수리', 29)).toBe(29);
    expect(scissorQty('복원수리(날변형)', 1)).toBe(1);
  });
  it('배송비 · 배송 · 각인은 0', () => {
    expect(scissorQty('배송비', 1)).toBe(0);
    expect(scissorQty('배송', 1)).toBe(0);
    expect(scissorQty('각인', 2)).toBe(0);
  });
  it('추가 작업 줄은 0', () => {
    expect(scissorQty(repairExtraName('날 변형 교정'), 1)).toBe(0);
    expect(isNonScissorRepairLine('추가작업 · 날 교체')).toBe(true);
  });
  it('추가 작업 이름 — 한 번만 붙고, 배송비는 그대로', () => {
    expect(repairExtraName('날 변형 교정')).toBe('추가작업 · 날 변형 교정');
    expect(repairExtraName('추가작업 · 날 변형 교정')).toBe('추가작업 · 날 변형 교정');
    expect(repairExtraName('배송비')).toBe('배송비');
  });
  it('이름이 비어도 터지지 않는다', () => {
    expect(scissorQty(null, 4)).toBe(4);
    expect(scissorQty(undefined, 0)).toBe(0);
  });
});
