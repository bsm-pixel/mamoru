import { describe, it, expect } from 'vitest';
import { matchSearch, toChosung } from '@/lib/utils/search-match';

// 검색 칸 규칙(2026-11-15): 대소문자 · 띄어쓰기 · 하이픈 무시, 전화 뒤 4자리, 초성
describe('검색 맞추기', () => {
  it('대소문자 · 띄어쓰기 · 하이픈을 무시한다', () => {
    expect(matchSearch('bl 600', '알파 블런트', 'BL-600')).toBe(true);
    expect(matchSearch('BL600', 'BL-600')).toBe(true);
    expect(matchSearch('bl700', 'BL-600')).toBe(false);
  });
  it('전화번호 뒤 4자리로 찾는다', () => {
    expect(matchSearch('5678', '시험고객', '010-1234-5678')).toBe(true);
    expect(matchSearch('01012345678', '010-1234-5678')).toBe(true);
  });
  it('초성으로 찾는다', () => {
    expect(toChosung('알파 블런트')).toBe('ㅇㅍㅂㄹㅌ');
    expect(matchSearch('ㅇㅍ', '알파 블런트')).toBe(true);
    expect(matchSearch('ㅂㄹㅌ', '알파 블런트')).toBe(true);
    expect(matchSearch('ㅌㄴ', '알파 블런트')).toBe(false);
  });
  it('검색어가 비면 모두 맞는다, 빈 값은 건너뛴다', () => {
    expect(matchSearch('', '무엇이든')).toBe(true);
    expect(matchSearch('가', null, undefined, '가위')).toBe(true);
  });
});
