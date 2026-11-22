import { describe, it, expect } from 'vitest';
import { contractImagePath } from '@/lib/contracts/image-path';

// 2026-11-22 계약서 이미지 — 공개 주소 대신 저장소 위치를 저장하고 열 때 1분짜리 주소로
describe('계약서 이미지 위치 읽기', () => {
  it('새 값(저장소 위치)은 그대로', () => {
    expect(contractImagePath('contracts/abc.png')).toBe('contracts/abc.png');
  });
  it('옛 공개 주소도 위치로 바꾼다', () => {
    expect(contractImagePath('https://x.supabase.co/storage/v1/object/public/contracts/contracts/abc.png')).toBe('contracts/abc.png');
  });
  it('비었거나 다른 곳 주소면 없음', () => {
    expect(contractImagePath('')).toBeNull();
    expect(contractImagePath(null)).toBeNull();
    expect(contractImagePath('https://example.com/a.png')).toBeNull();
  });
});
