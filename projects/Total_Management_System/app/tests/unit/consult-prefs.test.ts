import { describe, it, expect } from 'vitest';
import { prefDaysOf, prefTimesOf, prefText } from '@/lib/consultation/prefs';

// 2026-11-22 출장 요청의 희망 요일 · 선호 시간대 — 고르지 않았으면 「상관없음」
describe('희망 요일 · 선호 시간대 보여 주는 말', () => {
  it('고른 값은 그대로', () => {
    expect(prefDaysOf('월,수')).toEqual(['월', '수']);
    expect(prefTimesOf(['오전', '저녁'])).toEqual(['오전', '저녁']);
    expect(prefText(prefDaysOf(['월', '수']))).toBe('월, 수');
  });
  it('비었으면 「상관없음」', () => {
    for (const v of ['', null, undefined, [], ',']) {
      expect(prefDaysOf(v)).toEqual(['상관없음']);
      expect(prefTimesOf(v)).toEqual(['상관없음']);
    }
  });
  it('시간대 「무관」도 「상관없음」', () => {
    expect(prefTimesOf('무관')).toEqual(['상관없음']);
    expect(prefTimesOf(['오전', '무관'])).toEqual(['상관없음']);
  });
});
