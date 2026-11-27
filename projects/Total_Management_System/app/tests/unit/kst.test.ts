import { describe, it, expect, vi, afterEach } from 'vitest';
import { kstDate, kstDateDots, todayKST, addDays, addMonths, monthStartKST, daysSinceKST } from '@/lib/utils/kst';

/** 4차 7번 — 날짜는 한국 시간 기준 (서버가 UTC 여도 같은 값) */
describe('한국 시간 날짜', () => {
  afterEach(() => vi.useRealTimers());

  it('한국 15일 아침 8시(UTC 14일 23시)는 15일이다', () => {
    vi.useFakeTimers(); vi.setSystemTime(new Date('2026-11-14T23:00:00Z'));
    expect(todayKST()).toBe('2026-11-15');
    expect(new Date().toISOString().slice(0, 10)).toBe('2026-11-14'); // 예전 방식은 하루 앞
    expect(monthStartKST()).toBe('2026-11-01');
  });
  it('월말 밤: 한국 12월 1일 새벽은 12월이다', () => {
    vi.useFakeTimers(); vi.setSystemTime(new Date('2026-11-30T15:30:00Z'));
    expect(todayKST()).toBe('2026-12-01');
    expect(monthStartKST()).toBe('2026-12-01');
  });
  it('저장된 시각을 한국 날짜로', () => {
    expect(kstDate('2026-11-14T21:16:00+00:00')).toBe('2026-11-15');
    expect(kstDate('2026-11-14T10:00:00Z')).toBe('2026-11-14');
    expect(kstDate('엉뚱한 값')).toBe('');
  });
  it('날짜 더하기·빼기', () => {
    expect(addDays('2026-11-15', -30)).toBe('2026-10-16');
    expect(addDays('2026-12-31', 1)).toBe('2027-01-01');
    expect(addMonths('2026-11-15', -1)).toBe('2026-10-15');
    expect(addMonths('2026-03-31', -1)).toBe('2026-02-28');
  });
  it('마지막 거래 며칠 전 — 오늘 거래는 0일, 음수가 나오지 않는다', () => {
    vi.useFakeTimers(); vi.setSystemTime(new Date('2026-11-14T23:00:00Z')); // 한국 15일 08시
    expect(daysSinceKST('2026-11-15')).toBe(0);
    expect(daysSinceKST('2026-11-14')).toBe(1);
    expect(daysSinceKST('2026-10-16')).toBe(30);
  });
});

// 2026-11-27 고객 화면 날짜(수리내역서 입고일·출고일/전달일, 후기 주문일) — 서버 UTC로 하루 밀리던 것
describe('고객 화면 한국 날짜 kstDateDots', () => {
  it('오전 9시 전(한국) 시각도 그날 날짜', () => {
    // 한국 10-06 23:08 = UTC 10-06 14:08 → 10-06 / 한국 11-07 08:30 = UTC 11-06 23:30 → 11-07
    expect(kstDateDots('2026-10-06T14:08:00Z')).toBe('2026. 10. 6.');
    expect(kstDateDots('2026-11-06T23:30:00Z')).toBe('2026. 11. 7.');
  });
  it('값이 없거나 이상하면 빈 칸', () => {
    expect(kstDateDots(null)).toBe('');
    expect(kstDateDots(undefined)).toBe('');
    expect(kstDateDots('')).toBe('');
    expect(kstDateDots('엉뚱한 값')).toBe('');
  });
});
