/**
 * 아임웹 OpenAPI 토큰 — 단일 직렬화 모듈 (SSOT)
 *
 * 배경: 아임웹 refresh_token 은 rotation(1회용) 방식 — 갱신할 때마다 새 refresh_token 으로 교체된다.
 * 과거엔 갱신 로직이 3곳(client.getOpenApiToken / client.forceRefreshOpenApiToken /
 * script-api.getOpenApiToken)에 흩어져, 크론과 온디맨드가 동시에 같은 refresh_token 으로 갱신하면
 * 한쪽이 소비한 토큰을 다른 쪽이 다시 써 30170(유효하지 않은 리프레시 토큰)으로 체인이 끊겼다.
 * (2026-09-12 재고 연동 끊김 사고)
 *
 * 해결: 모든 갱신을 이 모듈로 모으고, DB 락(system_settings.imweb_openapi.refresh_lock)으로
 * 인스턴스 간 직렬화 + 락 획득 후 재확인(single-flight)으로 불필요한 중복 rotation 을 막는다.
 * 락 값 = 만료 ISO 문자열(사전식==시간순) → Postgres UPDATE ... WHERE value < now 로 원자적 CAS.
 */
import { createServiceClient } from '@/lib/supabase/server';

const OPENAPI_TOKEN_URL = 'https://openapi.imweb.me/oauth2/token';
const LOCK_KEY = 'imweb_openapi.refresh_lock';
const EPOCH = '1970-01-01T00:00:00.000Z';
const LOCK_TTL_MS = 30_000;             // 락 최대 보유 시간(갱신 1회 = HTTP 1회라 충분)
const USABLE_MS = 50 * 60 * 1000;       // access_token 발급 후 50분 이내면 유효(아임웹 만료 ~1시간)
const RECENT_ROTATE_MS = 5 * 60 * 1000; // 5분 내 이미 갱신됐으면 크론은 재-rotation 생략(체인 보호)
const LOCK_WAIT_MS = 10_000;            // 락 대기 총 한도
const CACHE_SAFETY_MS = 60_000;         // 인메모리 캐시 안전 여유

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type DbAny = any;
function db(): DbAny {
  return createServiceClient() as DbAny;
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** 프로세스(람다 인스턴스) 로컬 캐시 — access_token 은 rotation 과 무관하게 자체 만료(1h)까지 유효 */
let cached: { token: string; updatedAtMs: number } | null = null;

interface TokenRow { access: string | null; refresh: string | null; updatedAtMs: number }

async function readTokens(dbAny: DbAny): Promise<TokenRow> {
  const { data } = await dbAny
    .from('system_settings')
    .select('key, value')
    .in('key', ['imweb_openapi.access_token', 'imweb_openapi.refresh_token', 'imweb_openapi.token_updated_at']);
  const map: Record<string, string> = {};
  for (const s of (data || [])) map[s.key] = s.value;
  return {
    access: map['imweb_openapi.access_token'] || null,
    refresh: map['imweb_openapi.refresh_token'] || null,
    updatedAtMs: map['imweb_openapi.token_updated_at'] ? new Date(map['imweb_openapi.token_updated_at']).getTime() : 0,
  };
}

/** 락 행이 없으면 만들어 둔다(이미 있으면 건드리지 않음 = 보유 중인 락을 리셋하지 않음) */
async function ensureLockRow(dbAny: DbAny) {
  await dbAny.from('system_settings').upsert(
    { key: LOCK_KEY, value: EPOCH, updated_at: new Date().toISOString() },
    { onConflict: 'key', ignoreDuplicates: true },
  );
}

/** 락 획득 시도 — 성공 시 우리가 설정한 만료 ISO 반환, 실패 시 null (원자적 CAS) */
async function acquireLock(dbAny: DbAny): Promise<string | null> {
  const nowISO = new Date().toISOString();
  const expISO = new Date(Date.now() + LOCK_TTL_MS).toISOString();
  const { data } = await dbAny
    .from('system_settings')
    .update({ value: expISO, updated_at: nowISO })
    .eq('key', LOCK_KEY)
    .lt('value', nowISO)     // 비어있음(만료)일 때만 갱신됨 — ISO 문자열 사전식==시간순
    .select('key');
  return (Array.isArray(data) && data.length > 0) ? expISO : null;
}

/** 우리가 잡은 락만 해제(그 사이 남이 가져갔으면 no-op) */
async function releaseLock(dbAny: DbAny, expISO: string) {
  await dbAny
    .from('system_settings')
    .update({ value: EPOCH, updated_at: new Date().toISOString() })
    .eq('key', LOCK_KEY)
    .eq('value', expISO);
}

/** 실제 rotation — refresh_token 으로 새 access/refresh 발급 후 저장(새 refresh_token 먼저 확정) */
async function rotate(dbAny: DbAny, refresh: string): Promise<TokenRow> {
  const res = await fetch(OPENAPI_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      clientId: process.env.IMWEB_OPENAPI_KEY || '',
      clientSecret: process.env.IMWEB_OPENAPI_SECRET || '',
      grantType: 'refresh_token',
      refreshToken: refresh,
    }),
  });
  const data = await res.json();
  if (!res.ok || data.statusCode !== 200) {
    throw new Error(`아임웹 OpenAPI 토큰 갱신 실패: ${res.status} ${JSON.stringify(data)}`);
  }
  const newAccess = data.data.accessToken;
  const newRefresh = data.data.refreshToken;
  const now = new Date().toISOString();
  // 체인 유실 방지: 새 refresh_token 을 가장 먼저 확정 저장
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.refresh_token', value: newRefresh, updated_at: now }, { onConflict: 'key' });
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.access_token', value: newAccess, updated_at: now }, { onConflict: 'key' });
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.token_updated_at', value: now, updated_at: now }, { onConflict: 'key' });
  return { access: newAccess, refresh: newRefresh, updatedAtMs: new Date(now).getTime() };
}

/**
 * 유효한 access_token 확보(필요 시 락 아래에서 직렬화된 rotation).
 * - 'ondemand' : access_token 이 50분 지나 만료됐을 때만 갱신
 * - 'force'    : 크론 keep-alive — 5분 내 이미 갱신 안 됐으면 갱신
 */
async function ensureToken(mode: 'ondemand' | 'force'): Promise<TokenRow> {
  const dbAny = db();

  // 온디맨드 빠른 경로: 인메모리 캐시
  if (mode === 'ondemand' && cached && Date.now() - cached.updatedAtMs < USABLE_MS - CACHE_SAFETY_MS) {
    return { access: cached.token, refresh: null, updatedAtMs: cached.updatedAtMs };
  }

  const needRotate = (t: TokenRow) => {
    const age = Date.now() - t.updatedAtMs;
    if (mode === 'force') return age >= RECENT_ROTATE_MS;
    return !t.access || age >= USABLE_MS;
  };

  let cur = await readTokens(dbAny);
  if (!cur.access && !cur.refresh) {
    throw new Error('아임웹 OpenAPI 토큰이 없습니다. 설정 > 상품·재고에서 [아임웹 재연결]을 진행해주세요.');
  }
  if (!needRotate(cur)) {
    if (cur.access) cached = { token: cur.access, updatedAtMs: cur.updatedAtMs };
    return cur;
  }

  await ensureLockRow(dbAny);
  const deadline = Date.now() + LOCK_WAIT_MS;
  while (Date.now() < deadline) {
    const held = await acquireLock(dbAny);
    if (held) {
      try {
        cur = await readTokens(dbAny);            // 락 안에서 재확인
        if (!needRotate(cur)) {                    // 대기 중 다른 인스턴스가 이미 갱신
          if (cur.access) cached = { token: cur.access, updatedAtMs: cur.updatedAtMs };
          return cur;
        }
        if (!cur.refresh) {
          throw new Error('아임웹 OpenAPI refresh_token 없음 — [아임웹 재연결] 필요');
        }
        const next = await rotate(dbAny, cur.refresh);
        if (next.access) cached = { token: next.access, updatedAtMs: next.updatedAtMs };
        return next;
      } finally {
        await releaseLock(dbAny, held);
      }
    }
    // 락을 다른 인스턴스가 쥠 → 잠깐 대기 후 그쪽 갱신 결과 재사용 시도
    await sleep(400);
    cur = await readTokens(dbAny);
    if (!needRotate(cur) && cur.access) {
      cached = { token: cur.access, updatedAtMs: cur.updatedAtMs };
      return cur;
    }
  }
  throw new Error('아임웹 OpenAPI 토큰 갱신 락 대기 시간 초과 — 잠시 후 다시 시도해주세요.');
}

/** 유효한 OpenAPI access_token 반환(만료 시 직렬화된 rotation). 모든 호출처 공용. */
export async function getOpenApiToken(): Promise<string> {
  const t = await ensureToken('ondemand');
  if (!t.access) throw new Error('아임웹 OpenAPI access_token 확보 실패 — [아임웹 재연결] 필요');
  return t.access;
}

/**
 * 크론 keep-alive 강제 갱신. 만료 여부와 무관하게(단 5분 내 갱신됐으면 생략) rotation 체인을 살려둔다.
 * → 판매가 뜸한 기간에도 토큰이 방치로 죽지 않게 한다.
 */
export async function forceRefreshOpenApiToken(): Promise<{ ok: boolean; error?: string; updatedAt?: string }> {
  try {
    const t = await ensureToken('force');
    return { ok: true, updatedAt: new Date(t.updatedAtMs).toISOString() };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}
