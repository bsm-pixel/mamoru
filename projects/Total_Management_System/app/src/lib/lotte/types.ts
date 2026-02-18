/** 롯데택배 ALPS API 타입 */

export interface LotteConfig {
  url: string;
  cancelUrl: string;
  trackingUrl: string;
  clientKey: string;
  jobCustCd: string;
  sender: {
    name: string;
    tel: string;
    zip: string;
    addr: string;
  };
  fareSctCd: string;
}

export interface LotteBookRequest {
  ordNo: string;
  invNo?: string;
  rcvName: string;
  rcvTel: string;
  rcvZip: string;
  rcvAdr: string;
  boxTypCd?: string;
  gdsNm?: string;
  dlvMsg?: string;
  pickReqYmd?: string;
  ordSct?: '1' | '2' | '3'; // 1:일반 2:교환 3:AS (기본값 '1')
}

export interface LotteBookResult {
  ok: boolean;
  invNo: string;
  rtnCd: string;
  rtnMsg: string;
}

export interface LotteCancelResult {
  success: boolean;
  via?: string;
  error?: string;
}

export interface LotteTrackResult {
  ok: boolean;
  state: 'ACTIVE' | 'CANCELLED' | 'NOT_FOUND' | string;
  raw: Record<string, unknown>;
}
