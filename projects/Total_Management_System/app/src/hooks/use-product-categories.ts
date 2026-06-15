'use client';

/**
 * 상품 카테고리 옵션 SSOT — 설정(inventory.categories) 기반 + 시스템 고정 카테고리 보장.
 * 제품 등록/수정 드롭다운이 모두 이 훅을 사용 (이원화 제거, 2026-06-15 근본정리).
 */
import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CAT_LABELS, SYSTEM_CATEGORIES } from '@/lib/utils/setting-defaults';

export interface CategoryOption {
  value: string;
  label: string;
  system: boolean;
}

export function useProductCategoryOptions(): CategoryOption[] {
  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const categories = useSetting<string[]>('inventory.categories', Object.keys(DEFAULT_CAT_LABELS));

  const merged = [...categories];
  // 시스템 고정 카테고리(EVENT 등) — 설정에 없어도 항상 보장
  SYSTEM_CATEGORIES.forEach((c) => { if (!merged.includes(c)) merged.push(c); });

  return merged.map((c) => ({
    value: c,
    label: catLabels[c] || (c === 'EVENT' ? 'EVENT (재고 전환 이벤트)' : c),
    system: SYSTEM_CATEGORIES.includes(c),
  }));
}
