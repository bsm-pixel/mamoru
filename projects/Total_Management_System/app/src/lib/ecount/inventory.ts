/**
 * 이카운트 재고/품목 API
 * - 품목등록: /OAPI/V2/InventoryBasic/SaveBasicProduct
 * - 품목조회: /OAPI/V2/InventoryBasic/GetBasicProductsList
 * - 품목단건: /OAPI/V2/InventoryBasic/ViewBasicProduct
 * - 재고현황: /OAPI/V2/InventoryBalance/GetListInventoryBalanceStatus
 * - 재고단건: /OAPI/V2/InventoryBalance/ViewInventoryBalanceStatus
 */

import { ecountFetch, type EcountResponse } from './client';

/** 품목 조회 (목록) */
export async function listProducts(keyword?: string): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/InventoryBasic/GetBasicProductsList', {
    PROD_DES: keyword || '',
  });
}

/** 품목 단건 조회 */
export async function viewProduct(prodCode: string): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/InventoryBasic/ViewBasicProduct', {
    PROD_CD: prodCode,
  });
}

/** 품목 등록/수정 */
export async function saveProduct(params: {
  prodCode: string;
  prodName: string;
  unit?: string;
  remarks?: string;
}): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/InventoryBasic/SaveBasicProduct', {
    ProductList: [{
      BulkDatas: {
        PROD_CD: params.prodCode,
        PROD_DES: params.prodName,
        UNIT: params.unit || '',
        REMARKS: params.remarks || '',
      },
      Line: '0',
    }],
  });
}

/** 재고 현황 조회 (전체, BASE_DATE 필수) */
export async function getInventoryBalance(
  baseDate?: string, // YYYYMMDD (기본: 오늘)
  warehouseCode?: string,
  productCode?: string,
): Promise<EcountResponse> {
  const today = baseDate || new Date().toISOString().slice(0, 10).replace(/-/g, '');
  return ecountFetch('/OAPI/V2/InventoryBalance/GetListInventoryBalanceStatus', {
    BASE_DATE: today,
    WH_CD: warehouseCode || '',
    PROD_CD: productCode || '',
  });
}

/** 재고 현황 단건 조회 */
export async function viewInventoryBalance(
  productCode: string,
  baseDate?: string, // YYYYMMDD (기본: 오늘)
  warehouseCode?: string,
): Promise<EcountResponse> {
  const today = baseDate || new Date().toISOString().slice(0, 10).replace(/-/g, '');
  return ecountFetch('/OAPI/V2/InventoryBalance/ViewInventoryBalanceStatus', {
    BASE_DATE: today,
    PROD_CD: productCode,
    WH_CD: warehouseCode || '',
  });
}
