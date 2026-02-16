/**
 * 아임웹 API v2 클라이언트
 * 인증 흐름: POST /v2/auth → access_token 발급
 */

import type {
  ImwebAuthResponse,
  ImwebApiResponse,
  ImwebOrderListData,
  ImwebOrder,
  ImwebOrdersParams,
  ImwebProdOrder,
} from './types';

const BASE_URL = 'https://api.imweb.me';

let cachedToken: { token: string; expiresAt: number } | null = null;

/** 액세스 토큰 발급 (캐싱) */
async function getAccessToken(): Promise<string> {
  if (cachedToken && Date.now() < cachedToken.expiresAt - 60_000) {
    return cachedToken.token;
  }

  const res = await fetch(`${BASE_URL}/v2/auth`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      key: process.env.IMWEB_API_KEY,
      secret: process.env.IMWEB_API_SECRET,
    }),
  });

  if (!res.ok) {
    throw new Error(`아임웹 인증 실패: ${res.status}`);
  }

  const data: ImwebAuthResponse = await res.json();
  cachedToken = {
    token: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };
  return cachedToken.token;
}

/** 인증된 API 요청 */
async function imwebFetch<T>(
  path: string,
  options: RequestInit = {}
): Promise<ImwebApiResponse<T>> {
  const token = await getAccessToken();
  const res = await fetch(`${BASE_URL}${path}`, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      'access-token': token,
      ...options.headers,
    },
  });

  if (!res.ok) {
    throw new Error(`아임웹 API 오류: ${res.status} ${await res.text()}`);
  }

  return res.json();
}

/** 주문 목록 조회 */
export async function getOrders(
  params: ImwebOrdersParams = {}
): Promise<ImwebApiResponse<ImwebOrderListData>> {
  const query = new URLSearchParams();
  if (params.order_date_from) query.set('order_date_from', String(params.order_date_from));
  if (params.order_date_to) query.set('order_date_to', String(params.order_date_to));
  if (params.status) query.set('status', params.status);
  if (params.page) query.set('page', String(params.page));
  if (params.limit) query.set('limit', String(params.limit || 50));

  const qs = query.toString();
  return imwebFetch<ImwebOrderListData>(`/v2/shop/orders${qs ? '?' + qs : ''}`);
}

/** 주문 단건 조회 */
export async function getOrder(
  orderNo: string
): Promise<ImwebApiResponse<ImwebOrder>> {
  return imwebFetch<ImwebOrder>(`/v2/shop/orders/${orderNo}`);
}

/** 주문 품목(prod-orders) 조회 */
export async function getProdOrders(
  orderNo: string
): Promise<ImwebApiResponse<ImwebProdOrder[]>> {
  return imwebFetch<ImwebProdOrder[]>(`/v2/shop/orders/${orderNo}/prod-orders`);
}

/** 송장 업데이트 */
export async function updateInvoice(
  orderNo: string,
  data: { parcel_code: string; invoice_no: string }
): Promise<ImwebApiResponse<unknown>> {
  return imwebFetch(`/v2/shop/orders/${orderNo}/prod/0`, {
    method: 'PATCH',
    body: JSON.stringify({
      parcel_code: data.parcel_code,
      invoice_no: data.invoice_no,
    }),
  });
}
