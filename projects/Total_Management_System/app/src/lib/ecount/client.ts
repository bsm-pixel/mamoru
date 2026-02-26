/**
 * 이카운트 OAPI V2 클라이언트
 * - ZONE 조회 → 로그인(SESSION_ID 발급) → API 호출
 * - SESSION_ID는 서버 메모리 캐시 (만료 시 자동 재발급)
 */

const ECOUNT_BASE = 'https://sboapi.ecount.com';

interface EcountConfig {
  comCode: string;
  userId: string;
  apiCertKey: string;
  zone: string;
}

interface EcountResponse<T = unknown> {
  Status: string;   // '200' = 성공
  Error: { StatusCode: string; Message: string } | null;
  Data: T;
}

// 서버 메모리 캐시 (Vercel serverless = 짧은 수명이지만 연속 호출 시 유효)
let cachedSession: { id: string; expiresAt: number } | null = null;

function getConfig(): EcountConfig {
  const comCode = process.env.ECOUNT_COM_CODE;
  const userId = process.env.ECOUNT_USER_ID;
  const apiCertKey = process.env.ECOUNT_API_CERT_KEY;
  const zone = process.env.ECOUNT_ZONE;

  if (!comCode || !userId || !apiCertKey || !zone) {
    throw new Error('이카운트 환경변수 미설정: ECOUNT_COM_CODE, ECOUNT_USER_ID, ECOUNT_API_CERT_KEY, ECOUNT_ZONE');
  }
  return { comCode, userId, apiCertKey, zone };
}

function getBaseUrl(zone: string): string {
  return `https://sboapi${zone}.ecount.com`;
}

/** ZONE 조회 (초기 1회) */
export async function getZone(comCode: string): Promise<string> {
  const res = await fetch(`${ECOUNT_BASE}/OAPI/V2/Zone?COM_CODE=${comCode}`);
  const json = await res.json() as EcountResponse<{ ZONE: string }>;
  if (json.Status !== '200' || !json.Data?.ZONE) {
    throw new Error(`이카운트 ZONE 조회 실패: ${json.Error?.Message || 'Unknown'}`);
  }
  return json.Data.ZONE;
}

/** 로그인 → SESSION_ID 발급 */
async function login(): Promise<string> {
  const cfg = getConfig();
  const base = getBaseUrl(cfg.zone);

  const res = await fetch(`${base}/OAPI/V2/OAPILogin`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      COM_CODE: cfg.comCode,
      USER_ID: cfg.userId,
      API_CERT_KEY: cfg.apiCertKey,
      LAN_TYPE: 'ko-KR',
      ZONE: cfg.zone,
    }),
  });

  const json = await res.json() as EcountResponse<{ Datas: { SESSION_ID: string } }>;
  if (json.Status !== '200' || !json.Data?.Datas?.SESSION_ID) {
    throw new Error(`이카운트 로그인 실패: ${json.Error?.Message || JSON.stringify(json)}`);
  }
  return json.Data.Datas.SESSION_ID;
}

/** SESSION_ID 가져오기 (캐시 + 자동 재발급) */
export async function getSessionId(): Promise<string> {
  const now = Date.now();
  if (cachedSession && cachedSession.expiresAt > now) {
    return cachedSession.id;
  }
  const sessionId = await login();
  // 이카운트 세션은 약 1시간 유효 → 50분 캐시
  cachedSession = { id: sessionId, expiresAt: now + 50 * 60 * 1000 };
  return sessionId;
}

/** 이카운트 API 호출 래퍼 */
export async function ecountFetch<T = unknown>(
  path: string,
  body: unknown,
  retried = false,
): Promise<EcountResponse<T>> {
  const cfg = getConfig();
  const base = getBaseUrl(cfg.zone);
  const sessionId = await getSessionId();

  const res = await fetch(`${base}${path}?SESSION_ID=${sessionId}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });

  const json = await res.json() as EcountResponse<T>;

  // 세션 만료 시 1회 재발급 재시도
  if (json.Status !== '200' && json.Error?.StatusCode === '401' && !retried) {
    cachedSession = null;
    return ecountFetch<T>(path, body, true);
  }

  return json;
}

export type { EcountResponse, EcountConfig };
