import type { Product } from '@/lib/supabase/types';

/** 단가 그룹 정의 (settings에 저장) */
export interface PriceGroupDef {
  label: string;
  color: string;
  customerTypes: string[];
}

/** 기본 단가 그룹 (settings 로드 전 fallback) */
export const DEFAULT_PRICE_GROUPS: Record<string, PriceGroupDef> = {
  dealer: { label: '딜러가', color: 'purple', customerTypes: ['dealer'] },
  academy: { label: '아카데미가', color: 'emerald', customerTypes: ['academy'] },
};

/** 고객 유형 → 단가 그룹 키 resolve. 매칭 없으면 undefined */
export function resolvePriceGroup(
  customerType: string | undefined,
  groupDefs: Record<string, PriceGroupDef>,
): string | undefined {
  if (!customerType) return undefined;
  for (const [key, def] of Object.entries(groupDefs)) {
    if (def.customerTypes.includes(customerType)) return key;
  }
  return undefined;
}

/** 고객 유형에 따른 단가 결정. 그룹 가격이 없으면 소매가 fallback */
export function getUnitPrice(
  product: Product,
  customerType?: string,
  groupDefs?: Record<string, PriceGroupDef>,
): number {
  if (!customerType || !groupDefs) return product.price;
  const groupKey = resolvePriceGroup(customerType, groupDefs);
  if (!groupKey) return product.price;
  const groupPrice = product.price_groups?.[groupKey]?.price;
  return (groupPrice && groupPrice > 0) ? groupPrice : product.price;
}

/** 고객 유형에 따른 납품명 결정. 그룹 납품명이 없으면 제품명 fallback */
export function getProductDisplayName(
  product: Product,
  customerType?: string,
  groupDefs?: Record<string, PriceGroupDef>,
): string {
  if (!customerType || !groupDefs) return product.name;
  const groupKey = resolvePriceGroup(customerType, groupDefs);
  if (!groupKey) return product.name;
  const displayName = product.price_groups?.[groupKey]?.display_name;
  return displayName || product.name;
}

/** 해당 고객 유형에 대한 그룹 가격이 존재하는지 여부 */
export function hasGroupPrice(
  product: Product,
  customerType?: string,
  groupDefs?: Record<string, PriceGroupDef>,
): boolean {
  if (!customerType || !groupDefs) return false;
  const groupKey = resolvePriceGroup(customerType, groupDefs);
  if (!groupKey) return false;
  const groupPrice = product.price_groups?.[groupKey]?.price;
  return !!groupPrice && groupPrice > 0;
}
