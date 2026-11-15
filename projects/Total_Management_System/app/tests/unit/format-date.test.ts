import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { formatDate, formatDateFull, formatDateTime, formatDateTimeShort } from '@/lib/utils/format';

// 날짜 표기 통일(2026-11-15) — 목록 11.15 (일) / 상세 2026.11.15 (일) / 오늘·어제·내일은 글자 / 한국 시간 기준
describe('날짜 표기', () => {
  beforeEach(() => { vi.useFakeTimers(); vi.setSystemTime(new Date('2026-11-15T03:00:00Z')); }); // 한국 11/15 12:00
  afterEach(() => { vi.useRealTimers(); });

  it('목록: 오늘 · 어제 · 내일은 글자로, 그 밖은 월.일 (요일)', () => {
    expect(formatDate('2026-11-15')).toBe('오늘');
    expect(formatDate('2026-11-14')).toBe('어제');
    expect(formatDate('2026-11-16')).toBe('내일');
    expect(formatDate('2026-11-09')).toBe('11.09 (월)');
    expect(formatDate('2026-09-05')).toBe('09.05 (토)');
  });

  it('목록: 올해가 아니면 연도를 붙인다', () => {
    expect(formatDate('2025-11-15')).toBe('2025.11.15 (토)');
  });

  it('상세: 항상 연도까지, 글자로 바꾸지 않는다', () => {
    expect(formatDateFull('2026-11-15')).toBe('2026.11.15 (일)');
    expect(formatDateTime('2026-11-10T05:30:00Z')).toBe('2026.11.10 (화) 14:30');
  });

  it('시각이 있는 값은 한국 시간으로 날짜를 정한다 (UTC 로는 전날인 새벽)', () => {
    // UTC 11/8 16:30 = 한국 11/9 01:30
    expect(formatDate('2026-11-08T16:30:00Z')).toBe('11.09 (월)');
    expect(formatDateTimeShort('2026-11-08T16:30:00Z')).toBe('11.09 (월) 01:30');
    expect(formatDateTimeShort('2026-11-15T00:10:00Z')).toBe('오늘 09:10');
  });

  it('비었거나 잘못된 값은 빈 글자 (화면이 깨지지 않게)', () => {
    expect(formatDate('')).toBe('');
    expect(formatDate(null)).toBe('');
    expect(formatDateFull(undefined)).toBe('');
    expect(formatDateTime('없는날짜')).toBe('');
  });

  it('모양을 직접 주면 그대로', () => {
    expect(formatDate('2026-11-09', 'M.d')).toBe('11.9');
  });
});

import { formatPhone } from '@/lib/utils/format';
describe('전화번호 표기 — 보여 줄 때는 항상 하이픈', () => {
  it('휴대폰 · 서울 · 대표번호 · 안심번호', () => {
    expect(formatPhone('01012345678')).toBe('010-1234-5678');
    expect(formatPhone('010-1234-5678')).toBe('010-1234-5678');
    expect(formatPhone('0212345678')).toBe('02-1234-5678');
    expect(formatPhone('021234567')).toBe('02-123-4567');
    expect(formatPhone('0311234567')).toBe('031-123-4567');
    expect(formatPhone('15881234')).toBe('1588-1234');
    expect(formatPhone('050712345678')).toBe('0507-1234-5678');
  });
  it('해외 · 빈 값 · 알 수 없는 모양은 그대로', () => {
    expect(formatPhone('+81 90 1234 5678')).toBe('+81 90 1234 5678');
    expect(formatPhone(null)).toBe('');
    expect(formatPhone('1234')).toBe('1234');
  });
});
