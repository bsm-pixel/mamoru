/**
 * Kakao Local API 기반 주소 → 좌표 변환 공통 helper
 *
 * 회귀 방지를 위해 모든 consultation 좌표 채우기 경로(public/submit, admin-create,
 * backfill-coords)에서 이 함수만 사용하도록 통합. 한 곳 fix가 모든 경로에 자동 반영됨.
 *
 * 3단계 fallback (정확도 → 유연성):
 *   1) address.json(road)      — 카카오 등록된 도로명/지번 정확 매칭
 *   2) keyword.json(road)      — 신축/오래된 주소 유연 매칭
 *   3) keyword.json(road+detail) — 상호명·건물명 포함 케이스 (detail 있을 때)
 *
 * 빌라/오피스텔 상세주소(예: "OO빌딩 3층 301호")가 붙으면 1)이 매칭 실패하므로
 * 반드시 addressRoad 단독으로 1단계를 시도해야 한다. 합쳐서 던지면 회귀 발생.
 */

const KAKAO_REST_KEY = process.env.KAKAO_REST_API_KEY || '';
const TIMEOUT_MS = 3000;

export type GeocodeResult = { lat: number; lng: number; via: 'address' | 'keyword(road)' | 'keyword(road+detail)' };

async function tryEndpoint(path: '/v2/local/search/address.json' | '/v2/local/search/keyword.json', query: string): Promise<{ lat: number; lng: number } | null> {
  if (!KAKAO_REST_KEY || !query) return null;
  try {
    const res = await fetch(`https://dapi.kakao.com${path}?query=${encodeURIComponent(query)}`, {
      headers: { Authorization: `KakaoAK ${KAKAO_REST_KEY}` },
      signal: AbortSignal.timeout(TIMEOUT_MS),
    });
    if (!res.ok) return null;
    const json = await res.json();
    const doc = json.documents?.[0];
    if (doc?.y && doc?.x) return { lat: parseFloat(doc.y), lng: parseFloat(doc.x) };
  } catch { /* 다음 fallback */ }
  return null;
}

/**
 * 출장상담 주소를 좌표로 변환. 반드시 addressRoad만 1단계에 사용.
 * 매칭 실패 시 null 반환 — 호출 측은 NULL 허용으로 INSERT 진행해야 함.
 */
export async function geocodeForConsultation(
  addressRoad: string | null | undefined,
  addressDetail?: string | null,
): Promise<GeocodeResult | null> {
  const road = (addressRoad || '').trim();
  if (!road) return null;

  const a = await tryEndpoint('/v2/local/search/address.json', road);
  if (a) return { ...a, via: 'address' };

  const k = await tryEndpoint('/v2/local/search/keyword.json', road);
  if (k) return { ...k, via: 'keyword(road)' };

  const detail = (addressDetail || '').trim();
  if (detail) {
    const k2 = await tryEndpoint('/v2/local/search/keyword.json', `${road} ${detail}`);
    if (k2) return { ...k2, via: 'keyword(road+detail)' };
  }

  return null;
}
