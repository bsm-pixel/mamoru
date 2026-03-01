/**
 * 이카운트 거래처 API
 * - 거래처 등록: /OAPI/V2/AccountBasic/SaveBasicCust
 * - 거래처 조회: /OAPI/V2/AccountBasic/GetBasicCustList
 * - 거래처 단건: /OAPI/V2/AccountBasic/ViewBasicCust
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

/** 이카운트 거래처 레코드 */
export interface EcountCustomer {
  CUST_CD: string;         // 거래처코드
  CUST_DES: string;        // 거래처명
  TEL?: string;            // 전화번호
  EMAIL?: string;          // 이메일
  ADDR?: string;           // 주소
  REMARKS?: string;        // 비고
  [key: string]: unknown;
}

/** 거래처 목록 조회 */
export async function listCustomers(keyword?: string): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/AccountBasic/GetBasicCustList', {
    CUST_DES: keyword || '',
  });
}

/** 거래처 단건 조회 */
export async function viewCustomer(custCode: string): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/AccountBasic/ViewBasicCust', {
    CUST_CD: custCode,
  });
}

/** 거래처 등록/수정 */
export async function saveCustomer(params: SaveCustomerParams): Promise<EcountResponse> {
  return ecountFetch('/OAPI/V2/AccountBasic/SaveBasicCust', {
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

export type { SaveCustomerParams };
