/**
 * 아임웹 Script API 클라이언트
 * https://openapi.imweb.me/script
 *
 * 스크립트 주입 용도로만 사용 — widget.js 태그 자동 등록
 * 스펙:
 *   - GET    /script?unitCode=X&position=Y      (조회)
 *   - POST   /script  body: { unitCode, position, scriptContent }   (등록)
 *   - PUT    /script  body: { unitCode, position, scriptContent }   (수정)
 *   - DELETE /script?unitCode=X&position=Y      (삭제)
 *
 * 제약:
 *   - 동일 unitCode+position 조합 중복 POST 시 에러 30174
 *   - 따라서 GET 후 있으면 PUT, 없으면 POST
 */

import { createServiceClient } from '@/lib/supabase/server';

const OPENAPI_BASE = 'https://openapi.imweb.me';

export type ScriptPosition = 'header' | 'body' | 'footer';

export interface ImwebScriptRecord {
  siteCode: string;
  unitCode: string;
  position: ScriptPosition;
  scriptContent: string;
  wtime: string;
  mtime: string;
}

/** OpenAPI 토큰 로드 (만료 시 refresh) — client.ts의 로직 재사용 */
async function getOpenApiToken(): Promise<string> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const { data: settings } = await dbAny
    .from('system_settings')
    .select('key, value')
    .in('key', [
      'imweb_openapi.access_token',
      'imweb_openapi.refresh_token',
      'imweb_openapi.token_updated_at',
    ]);

  if (!settings || settings.length === 0) {
    throw new Error('아임웹 OpenAPI 토큰이 없습니다. OAuth 재연결이 필요합니다.');
  }

  const map: Record<string, string> = {};
  for (const s of settings) map[s.key] = s.value;

  const accessToken = map['imweb_openapi.access_token'];
  const refreshToken = map['imweb_openapi.refresh_token'];
  const updatedAt = map['imweb_openapi.token_updated_at'];

  if (!accessToken) {
    throw new Error('아임웹 OpenAPI access_token 없음 — OAuth 재연결 필요');
  }

  // 50분 이내면 그대로 사용
  const tokenAge = Date.now() - new Date(updatedAt || 0).getTime();
  if (tokenAge < 50 * 60 * 1000) {
    return accessToken;
  }

  // 만료 → refresh
  if (!refreshToken) {
    throw new Error('refresh_token 없음 — OAuth 재연결 필요');
  }

  const res = await fetch(`${OPENAPI_BASE}/oauth2/token`, {
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
    throw new Error(`토큰 갱신 실패: ${res.status} ${JSON.stringify(data)}`);
  }

  const newAccess = data.data.accessToken;
  const newRefresh = data.data.refreshToken;
  const now = new Date().toISOString();

  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.access_token', value: newAccess, updated_at: now }, { onConflict: 'key' });
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.refresh_token', value: newRefresh, updated_at: now }, { onConflict: 'key' });
  await dbAny.from('system_settings').upsert({ key: 'imweb_openapi.token_updated_at', value: now, updated_at: now }, { onConflict: 'key' });

  return newAccess;
}

interface ImwebApiEnvelope<T> {
  statusCode: number;
  data?: T;
  error?: { errorCode: string; message: string };
}

/** 공통 fetch 래퍼 */
async function scriptApiFetch<T>(
  method: 'GET' | 'POST' | 'PUT' | 'DELETE',
  path: string,
  body?: Record<string, unknown>,
): Promise<ImwebApiEnvelope<T>> {
  const token = await getOpenApiToken();

  const res = await fetch(`${OPENAPI_BASE}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(body ? { 'Content-Type': 'application/json' } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  const json = (await res.json().catch(() => ({}))) as ImwebApiEnvelope<T>;
  return json;
}

/** GET /script — 해당 unitCode/position의 스크립트 조회 */
export async function getScript(
  unitCode: string,
  position?: ScriptPosition,
): Promise<ImwebScriptRecord[]> {
  const query = new URLSearchParams({ unitCode });
  if (position) query.set('position', position);
  const res = await scriptApiFetch<ImwebScriptRecord[]>('GET', `/script?${query}`);
  if (res.statusCode !== 200) {
    throw new Error(`Script API GET 실패: ${JSON.stringify(res)}`);
  }
  return res.data || [];
}

/** POST /script — 신규 등록 */
export async function createScript(
  unitCode: string,
  position: ScriptPosition,
  scriptContent: string,
): Promise<void> {
  const res = await scriptApiFetch<boolean>('POST', '/script', {
    unitCode,
    position,
    scriptContent,
  });
  if (res.statusCode !== 200) {
    throw new Error(`Script API POST 실패: ${JSON.stringify(res)}`);
  }
}

/** PUT /script — 기존 스크립트 교체 */
export async function updateScript(
  unitCode: string,
  position: ScriptPosition,
  scriptContent: string,
): Promise<void> {
  const res = await scriptApiFetch<boolean>('PUT', '/script', {
    unitCode,
    position,
    scriptContent,
  });
  if (res.statusCode !== 200) {
    throw new Error(`Script API PUT 실패: ${JSON.stringify(res)}`);
  }
}

/** DELETE /script */
export async function deleteScript(
  unitCode: string,
  position: ScriptPosition,
): Promise<void> {
  const query = new URLSearchParams({ unitCode, position });
  const res = await scriptApiFetch<boolean>('DELETE', `/script?${query}`);
  if (res.statusCode !== 200) {
    throw new Error(`Script API DELETE 실패: ${JSON.stringify(res)}`);
  }
}

/**
 * 우리 widget.js를 아임웹에 업서트(있으면 PUT, 없으면 POST)
 *
 * @param unitCode 아임웹 유닛 코드
 * @param position 주입 위치 (기본 'footer' — defer 효과)
 * @param widgetUrl 전체 URL (예: https://app.../api/imweb/banner-widget.js)
 */
export async function upsertMamoruWidget(
  unitCode: string,
  position: ScriptPosition,
  widgetUrl: string,
): Promise<{ action: 'created' | 'updated' }> {
  const scriptContent = `<script src="${widgetUrl}" defer></script>`;

  // 이미 등록된 것이 있는지 조회
  const existing = await getScript(unitCode, position);
  const already = existing.find((s) => s.position === position);

  if (already) {
    // 내용이 동일하면 스킵, 다르면 PUT
    if (already.scriptContent.trim() === scriptContent.trim()) {
      return { action: 'updated' }; // 이미 최신이지만 UX상 'updated' 취급
    }
    await updateScript(unitCode, position, scriptContent);
    return { action: 'updated' };
  }

  await createScript(unitCode, position, scriptContent);
  return { action: 'created' };
}
