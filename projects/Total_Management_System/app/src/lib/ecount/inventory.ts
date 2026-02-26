/**
 * 이카운트 재고 API
 * - 재고 현황 조회 (GetInventoryOnhand)
 */

import { ecountFetch, type EcountResponse } from './client';

/** 재고 현황 조회 */
export async function getInventoryOnhand(
  warehouseCode?: string,
  productCode?: string,
): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/Inventory/GetListInventoryOnhandSp', {
    WH_CD: warehouseCode || '',
    PROD_CD: productCode || '',
  });
}

/** 품목 조회 */
export async function listProducts(keyword?: string): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/InventoryBasic/GetListBasicProdSp', {
    PROD_DES: keyword || '',
  });
}
