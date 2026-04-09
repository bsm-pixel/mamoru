import { useSetting } from './use-settings';
import { DEFAULT_PRICE_GROUPS, type PriceGroupDef } from '@/lib/utils/pricing';

/** settings의 pricing.groups를 읽어 단가 그룹 정의를 반환 */
export function usePriceGroups(): Record<string, PriceGroupDef> {
  return useSetting<Record<string, PriceGroupDef>>('pricing.groups', DEFAULT_PRICE_GROUPS);
}
