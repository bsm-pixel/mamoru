/**
 * 아임웹 API 클라이언트
 * v2 API: api.imweb.me (주문 관리)
 * 새 OpenAPI: openapi.imweb.me (상품/재고 관리)
 * 인증 흐름: POST /v2/auth → access_token 발급 (양쪽 공용)
 */

import type {
  ImwebAuthResponse,
  ImwebApiResponse,
  ImwebOrderListData,
  ImwebOrder,
  ImwebOrdersParams,
  ImwebProdOrder,
  ImwebProduct,
  ImwebProductListResponse,
} from './types';

const BASE_URL = 'https://api.imweb.me';
const OPENAPI_URL = 'https://openapi.imweb.me';

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

/** 새 OpenAPI 인증된 요청 (Bearer 토큰) */
async function openapiFetch<T>(
  path: string,
  options: RequestInit = {}
): Promise<T> {
  const token = await getAccessToken();
  const res = await fetch(`${OPENAPI_URL}${path}`, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
      ...options.headers,
    },
  });

  if (!res.ok) {
    throw new Error(`아임웹 OpenAPI 오류: ${res.status} ${await res.text()}`);
  }

  return res.json();
}

/* ============================================================
 * 상품 API (새 OpenAPI — openapi.imweb.me)
 * ============================================================ */

/** 상품 목록 조회 */
export async function getImwebProducts(
  page = 1,
  limit = 100
): Promise<ImwebProductListResponse> {
  return openapiFetch<ImwebProductListResponse>(
    `/products?page=${page}&limit=${limit}`
  );
}

/** 상품 단건 조회 */
export async function getImwebProduct(
  prodNo: number
): Promise<{ statusCode: number; data: ImwebProduct }> {
  return openapiFetch(`/products/${prodNo}`);
}

/** 상품 재고 수정 (+ or -) */
export async function updateImwebStock(
  prodNo: number,
  quantity: number
): Promise<{ statusCode: number; data: boolean }> {
  return openapiFetch(`/products/${prodNo}/stock-info`, {
    method: 'PATCH',
    body: JSON.stringify({
      stock: quantity,
      isUseStock: 'Y',
      isUnlimitedStock: 'N',
    }),
  });
}

/* ============================================================
 * 주문 API (v2 — api.imweb.me)
 * ============================================================ */

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

/**
 * 아임웹 prod-order 상태 조회
 * 송장 입력 가능 여부 사전 체크용
 */
export async function getImwebOrderStatus(
  orderNo: string
): Promise<{ prodOrderNo: string | null; status: string }> {
  const prodRes = await getProdOrders(orderNo);
  const prodOrders = prodRes.data || [];
  if (prodOrders.length === 0) return { prodOrderNo: null, status: 'unknown' };
  return { prodOrderNo: prodOrders[0].order_no, status: prodOrders[0].status };
}

/** 송장 입력이 가능한 아임웹 상태 */
const INVOICE_READY_STATUSES = ['STANDBY', 'DELIVERY_READY', 'DELIVERY', 'DELIVERING'];

/**
 * 송장 등록 — 아임웹 배송대기 이상일 때만 동작
 * @returns success: 아임웹 반영 성공 여부, needsManual: 수동 처리 필요 여부
 */
export async function updateInvoice(
  orderNo: string,
  data: { parcel_code: string; invoice_no: string }
): Promise<{ success: boolean; needsManual: boolean; imwebStatus: string }> {
  // 1) 현재 아임웹 상태 확인
  const { prodOrderNo, status } = await getImwebOrderStatus(orderNo);
  console.log('[imweb/updateInvoice] 상태:', { orderNo, prodOrderNo, status });

  if (!prodOrderNo) {
    return { success: false, needsManual: true, imwebStatus: 'no_prod_order' };
  }

  // 2) 배송대기 이상이 아니면 → 수동 처리 필요
  if (!INVOICE_READY_STATUSES.includes(status)) {
    console.log('[imweb/updateInvoice] 아임웹 배송대기 미전환 → 수동 필요');
    return { success: false, needsManual: true, imwebStatus: status };
  }

  // 3) 송장번호 입력 (공식 엔드포인트)
  const res = await imwebFetch<unknown>(
    `/v2/shop/prod-orders/${prodOrderNo}/invoice`,
    {
      method: 'PATCH',
      body: JSON.stringify({
        parcel_code: data.parcel_code,
        invoice_no: data.invoice_no,
        order_version: 'v2',
      }),
    }
  );

  const ok = res.code === 200 || res.code === 0;
  console.log('[imweb/updateInvoice] 송장 입력:', { code: res.code, msg: res.msg, ok });
  return { success: ok, needsManual: !ok, imwebStatus: status };
}

/** 품목주문 상태 변경 (CANCEL 등) */
export async function updateProdOrderStatus(
  orderNo: string,
  status: string
): Promise<ImwebApiResponse<unknown>> {
  const prodRes = await getProdOrders(orderNo);
  const prodOrders = prodRes.data || [];

  if (prodOrders.length > 0) {
    const prodOrderNo = prodOrders[0].order_no;
    return imwebFetch(`/v2/shop/prod-orders/${prodOrderNo}`, {
      method: 'PATCH',
      body: JSON.stringify({ status, order_version: 'v2' }),
    });
  }

  return imwebFetch(`/v2/shop/orders/${orderNo}/prod/0`, {
    method: 'PATCH',
    body: JSON.stringify({ status }),
  });
}
