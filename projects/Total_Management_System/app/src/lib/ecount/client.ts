/**
 * 이카운트 OAPI V2 클라이언트
 * - 테스트 인증키: https://sboapi{ZONE}.ecount.com
 * - 정식 인증키: https://oapi{ZONE}.ecount.com
 * - 로그인(SESSION_ID 발급) → API 호출
 * - SESSION_ID는 서버 메모리 캐시 (만료 시 자동 재발급)
 */

// ECOUNT_TEST_MODE=true → sboapi (테스트키), false/미설정 → oapi (정식키)
const isTestMode = process.env.ECOUNT_TEST_MODE === 'true';

interface EcountConfig {
  comCode: string;
  userId: string;
  apiCertKey: string;
  zone: string;
}

interface EcountResponse<T = unknown> {
  Status: string | number;
  Error: { Code: number; Message: string; MessageDetail: string; XErrors: unknown } | null;
  Errors?: Array<{ ProgramId: string; Name: string; Code: string; Message: string; Param: unknown }>;
  Data: T;
}

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
  const prefix = isTestMode ? 'sboapi' : 'oapi';
  return `https://${prefix}${zone}.ecount.com`;
}

/** Status 200 체크 (숫자/문자열 호환) */
function isStatusOk(status: string | number): boolean {
  return String(status) === '200';
}

/** ZONE 조회 (환경변수 검증용) */
export async function getZone(comCode: string): Promise<string> {
  const prefix = isTestMode ? 'sboapi' : 'oapi';
  const res = await fetch(`https://${prefix}.ecount.com/OAPI/V2/Zone?COM_CODE=${comCode}`);
  const json = await res.json() as EcountResponse<{ ZONE: string }>;
  if (!isStatusOk(json.Status) || !json.Data?.ZONE) {
    throw new Error(`ZONE 조회 실패: ${json.Error?.Message || JSON.stringify(json)}`);
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

  const json = await res.json() as EcountResponse<{ Datas: { SESSION_ID: string }; Code?: string; Message?: string }>;
  if (!isStatusOk(json.Status) || !json.Data?.Datas?.SESSION_ID) {
    // 테스트키를 oapi에 사용하면 Code 204 + "테스트용 인증키" 반환
    const hint = json.Data?.Code === '204'
      ? ' (테스트 인증키 → ECOUNT_TEST_MODE=true 필요)'
      : '';
    throw new Error(`이카운트 로그인 실패${hint}: ${json.Data?.Message || json.Error?.Message || JSON.stringify(json)}`);
  }
  return json.Data.Datas.SESSION_ID;
}

/** SESSION_ID (캐시 + 자동 재발급) */
export async function getSessionId(): Promise<string> {
  const now = Date.now();
  if (cachedSession && cachedSession.expiresAt > now) {
    return cachedSession.id;
  }
  const sessionId = await login();
  cachedSession = { id: sessionId, expiresAt: now + 50 * 60 * 1000 };
  return sessionId;
}

/** API 호출 래퍼 */
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
  if (!isStatusOk(json.Status) && json.Error?.Message?.includes?.('SESSION') && !retried) {
    cachedSession = null;
    return ecountFetch<T>(path, body, true);
  }

  return json;
}

export { isStatusOk };
export type { EcountResponse, EcountConfig };
