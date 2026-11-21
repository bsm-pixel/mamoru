import { describe, it, expect } from 'vitest';
import { pageSizeFor } from '@/hooks/use-page-size';

// 한 쪽 건수(2026-11-21) — 세로로 긴 화면만 늘리고, 보통 화면 · 휴대폰은 20건 그대로
describe('화면 높이에 맞는 건수', () => {
  it('보통 화면과 휴대폰은 20건', () => {
    expect(pageSizeFor(900, 1440)).toBe(20);
    expect(pageSizeFor(1080, 1920)).toBe(20);
    expect(pageSizeFor(2000, 390)).toBe(20);
  });
  it('세로로 긴 화면은 5건 단위로 늘고 60건을 넘지 않는다', () => {
    expect(pageSizeFor(1440, 2560)).toBe(25);
    expect(pageSizeFor(1920, 1080)).toBe(35);
    expect(pageSizeFor(2560, 1440)).toBe(50);
    expect(pageSizeFor(5000, 1440)).toBe(60);
  });
});
