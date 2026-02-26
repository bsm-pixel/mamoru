/**
 * 이카운트 거래처 API
 * - 거래처 등록 (SaveCustomer)
 * - 거래처 조회 (ListCustomerSp)
 */

import { ecountFetch, type EcountResponse } from './client';

interface SaveCustomerParams {
  custCode: string;        // 거래처코드 (자동 생성 or 지정)
  custName: string;        // 거래처명
  phone?: string;          // 전화번호
  email?: string;          // 이메일
  address?: string;        // 주소
  remarks?: string;        // 비고
}

/** 거래처 등록/수정 */
export async function saveCustomer(params: SaveCustomerParams): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/InventoryBasic/SaveBasicCust', {
    CustList: [{
      BulkDatas: {
        CUST_CD: params.custCode,
        CUST_DES: params.custName,
        TEL: params.phone || '',
        EMAIL: params.email || '',
        ADDR: params.address || '',
        REMARKS: params.remarks || '',
      },
      Line: '0',
    }],
  });
}

/** 거래처 조회 */
export async function listCustomers(keyword?: string): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/InventoryBasic/GetListBasicCustSp', {
    CUST_DES: keyword || '',
  });
}
