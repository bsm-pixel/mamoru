/**
 * 이카운트 판매 전표 API
 * - 판매입력: /OAPI/V2/Sale/SaveSale
 * - 판매 조회 API는 이카운트에서 미제공 (TMS Supabase에서 관리)
 */

import { ecountFetch, type EcountResponse } from './client';

interface SaleLineItem {
  PROD_CD: string;      // 품목코드
  PROD_DES?: string;    // 품목명
  QTY: string;          // 수량
  PRICE?: string;       // 단가
  SUPPLY_AMT?: string;  // 공급가
  VAT_AMT?: string;     // 부가세
  CUST_CD?: string;     // 거래처코드
  WH_CD?: string;       // 창고코드
  REMARKS?: string;     // 비고
}

interface SaveSaleParams {
  saleDate: string;           // YYYY-MM-DD
  customerCode?: string;      // 이카운트 거래처 코드
  warehouseCode?: string;     // 창고코드 (기본 '100')
  items: SaleLineItem[];
  remarks?: string;
}

/** 판매전표 생성 */
export async function saveSale(params: SaveSaleParams): Promise<EcountResponse> {
  const { saleDate, customerCode, warehouseCode = '100', items, remarks } = params;

  const SaleList = items.map((item, idx) => ({
    BulkDatas: {
      IO_DATE: saleDate.replace(/-/g, ''),  // YYYYMMDD
      CUST_CD: customerCode || '',
      WH_CD: item.WH_CD || warehouseCode,
      PROD_CD: item.PROD_CD,
      PROD_DES: item.PROD_DES || '',
      QTY: item.QTY,
      PRICE: item.PRICE || '',
      SUPPLY_AMT: item.SUPPLY_AMT || '',
      VAT_AMT: item.VAT_AMT || '0',
      REMARKS: item.REMARKS || remarks || '',
    },
    Line: String(idx),
  }));

  return ecountFetch('/OAPI/V2/Sale/SaveSale', { SaleList });
}

export type { SaveSaleParams, SaleLineItem };
