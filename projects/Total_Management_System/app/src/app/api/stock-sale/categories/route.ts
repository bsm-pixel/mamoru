import { NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/**
 * 재고판매(LS) 분류 목록 — 2026-07-21
 * 제품 등록 시 datalist 로 기존 분류를 재사용하기 위한 것.
 * LS 제품들의 tags.group 중 실제 쓰이는 값(중복 제거·이름순).
 * '분류 만들기' = 새 이름 타이핑, '삭제' = 그 분류의 제품을 다른 데로 옮기면 자동 소멸.
 */
export async function GET() {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { data, error } = await db
      .from('products')
      .select('tags')
      .eq('category', 'LS')
      .eq('is_active', true);
    if (error) throw error;

    const set = new Set<string>();
    for (const row of data || []) {
      const g = (row?.tags?.group ?? '').toString().trim();
      if (g) set.add(g);
    }
    return NextResponse.json({ categories: [...set].sort((a, b) => a.localeCompare(b, 'ko')) });
  } catch (err) {
    console.error('[stock-sale/categories] 실패:', err);
    return NextResponse.json({ categories: [], error: String(err) }, { status: 500 });
  }
}
