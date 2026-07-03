/**
 * 아임웹 API v2 클라이언트
 * 인증 흐름: POST /v2/auth → access_token 발급
 * 주문 + 상품 + 재고 모두 v2 API로 처리
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

/* ============================================================
 * 상품 API (v2)
 * ============================================================ */

/** 상품 목록 조회 */
export async function getImwebProducts(
  page = 1,
  limit = 50
): Promise<ImwebApiResponse<{ list: ImwebV2Product[]; pagenation: { data_count: string; current_page: number; total_page: number; pagesize: number } }>> {
  return imwebFetch(`/v2/shop/products?page=${page}&limit=${limit}`);
}

/** 새 OpenAPI 토큰 — DB에서 읽고 만료 시 refreshToken으로 갱신 */
let cachedOpenApiToken: { token: string; expiresAt: number } | null = null;

async function getOpenApiToken(): Promise<string> {
  // 메모리 캐시 유효하면 바로 반환
  if (cachedOpenApiToken && Date.now() < cachedOpenApiToken.expiresAt - 60_000) {
    return cachedOpenApiToken.token;
  }

  // DB에서 토큰 읽기
  const { createServiceClient } = await import('@/lib/supabase/server');
  const db = createServiceClient() as ReturnType<typeof createServiceClient> & { from: (...args: unknown[]) => unknown };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const { data: settings } = await dbAny
    .from('system_settings')
    .select('key, value')
    .in('key', ['imweb_openapi.access_token', 'imweb_openapi.refresh_token', 'imweb_openapi.token_updated_at']);

  if (!settings || settings.length === 0) {
    throw new Error('아임웹 OpenAPI 토큰이 없습니다. 설정 > 아임웹 연동에서 OAuth 인증을 진행해주세요.');
  }

  const tokenMap: Record<string, string> = {};
  for (const s of settings) tokenMap[s.key] = s.value;

  const accessToken = tokenMap['imweb_openapi.access_token'];
  const refreshToken = tokenMap['imweb_openapi.refresh_token'];
  const updatedAt = tokenMap['imweb_openapi.token_updated_at'];

  if (!accessToken) {
    throw new Error('아임웹 OpenAPI 토큰이 없습니다. OAuth 인증을 진행해주세요.');
  }

  // 토큰 발급 후 50분 이내면 유효 (아임웹 토큰 만료: 보통 1시간)
  const tokenAge = Date.now() - new Date(updatedAt || 0).getTime();
  if (tokenAge < 50 * 60 * 1000) {
    cachedOpenApiToken = { token: accessToken, expiresAt: new Date(updatedAt).getTime() + 60 * 60 * 1000 };
    return accessToken;
  }

  // 토큰 만료 → refreshToken으로 갱신
  if (!refreshToken) {
    throw new Error('아임웹 OpenAPI refreshToken이 없습니다. 재인증이 필요합니다.');
  }

  const res = await fetch('https://openapi.imweb.me/oauth2/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      clientId: process.env.IMWEB_OPENAPI_KEY || '',
      clientSecret: process.env.IMWEB_OPENAPI_SECRET || '',
      grantType: 'refresh_token',
      refreshToken,
    }),
  });

  const data = await res.json();

  if (!res.ok || data.statusCode !== 200) {
    throw new Error(`아임웹 OpenAPI 토큰 갱신 실패: ${res.status} ${JSON.stringify(data)}`);
  }

  const newAccess = data.data.accessToken;
  const newRefresh = data.data.refreshToken;
  const now = new Date().toISOString();

  // DB 업데이트
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.access_token', value: newAccess, updated_at: now }, { onConflict: 'key' });
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.refresh_token', value: newRefresh, updated_at: now }, { onConflict: 'key' });
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.token_updated_at', value: now, updated_at: now }, { onConflict: 'key' });

  cachedOpenApiToken = { token: newAccess, expiresAt: Date.now() + 60 * 60 * 1000 };
  return newAccess;
}

/**
 * 아임웹 OpenAPI 연결 상태 진단 (설정 화면용).
 * 재고 push가 의존하는 getOpenApiToken()을 실제로 호출해 유효성을 검증한다.
 *  - connected=false + error → 재인증(authorize) 필요 (refresh_token 만료 등)
 */
export async function getOpenApiConnectionStatus(): Promise<{
  connected: boolean;
  hasRefreshToken: boolean;
  updatedAt: string | null;
  ageMinutes: number | null;
  scope: string | null;
  error: string | null;
}> {
  const { createServiceClient } = await import('@/lib/supabase/server');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = createServiceClient() as any;
  const { data: settings } = await dbAny
    .from('system_settings')
    .select('key, value')
    .in('key', [
      'imweb_openapi.access_token',
      'imweb_openapi.refresh_token',
      'imweb_openapi.token_updated_at',
      'imweb_openapi.scope',
    ]);
  const map: Record<string, string> = {};
  for (const s of (settings || [])) map[s.key] = s.value;
  const updatedAt = map['imweb_openapi.token_updated_at'] || null;
  const scope = map['imweb_openapi.scope'] || null;
  const hasRefreshToken = !!map['imweb_openapi.refresh_token'];
  const ageMinutes = updatedAt ? Math.round((Date.now() - new Date(updatedAt).getTime()) / 60000) : null;
  try {
    await getOpenApiToken(); // 만료 시 refresh 시도 — 실패하면 여기서 throw
    return { connected: true, hasRefreshToken, updatedAt, ageMinutes, scope, error: null };
  } catch (e) {
    return { connected: false, hasRefreshToken, updatedAt, ageMinutes, scope, error: e instanceof Error ? e.message : String(e) };
  }
}

/**
 * 아임웹 재고 증감 — 새 OpenAPI 사용
 * @param delta 양수면 증가, 음수면 감소 (절대값 아님!)
 */
export async function adjustImwebStockDelta(
  prodNo: number,
  delta: number
): Promise<Record<string, unknown>> {
  if (delta === 0) return { skipped: true };
  const token = await getOpenApiToken();

  const res = await fetch(`https://openapi.imweb.me/products/${prodNo}/stock-info`, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`,
    },
    body: JSON.stringify({
      stock: delta,
      isUseStock: 'Y',
      isUnlimitedStock: 'N',
    }),
  });

  const data = await res.json();

  if (!res.ok) {
    throw new Error(`아임웹 OpenAPI 재고 수정 실패: ${res.status} ${JSON.stringify(data)}`);
  }

  return data;
}

/** 하위 호환 — 기존 호출처에서 사용. delta를 직접 전달 */
export const updateImwebStock = adjustImwebStockDelta;

/** v2 상품 응답 타입 */
export interface ImwebV2Product {
  no: number;
  name: string;
  price: number;
  price_org?: number;
  prod_status: string;
  custom_prod_code: string | null;
  categories: string[];
  images: string[];
  image_url: Record<string, string>;
  stock: {
    stock_use: boolean;
    stock_unlimit: boolean;
    stock_no_option: number;
    sku_no_option: string;
  };
  origin: string;
  maker: string;
  brand: string;
}

/* ============================================================
 * 주문 API (v2)
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
