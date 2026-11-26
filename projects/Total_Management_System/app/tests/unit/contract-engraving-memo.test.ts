import { describe, it, expect } from 'vitest';
import { applyEngravingToMemo } from '@/lib/contracts/engraving-memo';

/**
 * 계약서 각인 문구 → 판매 메모 (2026-11-26, 작업실 13번)
 * 판매로 바꾼 뒤 받은 각인 문구가 출고 준비표 · 판매 상세에 반영되는가
 */
describe('판매 메모의 각인 부분 맞추기', () => {
  it('「각인 문구 대기」 → 받은 문구', () => {
    expect(applyEngravingToMemo('계약서 CT-2611-001 · 각인 문구 대기', { on: true, text: 'K.M' })).toBe('계약서 CT-2611-001 · 각인 「K.M」');
  });
  it('문구를 고치면 옛 문구만 바뀐다(다른 메모는 그대로)', () => {
    expect(applyEngravingToMemo('계약서 CT-1 · 각인 「A」 · 택배 오전', { on: true, text: 'B' })).toBe('계약서 CT-1 · 각인 「B」 · 택배 오전');
  });
  it('문구를 비우면 다시 「각인 문구 대기」', () => {
    expect(applyEngravingToMemo('계약서 CT-1 · 각인 「A」', { on: true, text: null })).toBe('계약서 CT-1 · 각인 문구 대기');
  });
  it('각인을 끄면 각인 부분만 빠진다', () => {
    expect(applyEngravingToMemo('계약서 CT-1 · 각인 「A」 · 택배 오전', { on: false, text: null })).toBe('계약서 CT-1 · 택배 오전');
    expect(applyEngravingToMemo('각인 문구 대기', { on: false, text: null })).toBe(null);
  });
  it('메모에 각인 부분이 없으면 끝에 붙인다', () => {
    expect(applyEngravingToMemo('계약서 CT-1', { on: true, text: 'A' })).toBe('계약서 CT-1 · 각인 「A」');
    expect(applyEngravingToMemo(null, { on: true, text: 'A' })).toBe('각인 「A」');
  });
  it('각인 안 하는 계약 + 각인 부분 없는 메모는 그대로', () => {
    expect(applyEngravingToMemo('직접 입력 메모', { on: false, text: null })).toBe('직접 입력 메모');
    expect(applyEngravingToMemo(null, { on: false, text: null })).toBe(null);
  });
});
