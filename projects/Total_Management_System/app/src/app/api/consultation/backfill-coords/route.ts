import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

const KAKAO_REST_KEY = process.env.KAKAO_REST_API_KEY || '';

/** POST /api/consultation/backfill-coords — 좌표 없는 출장 건에 좌표 채우기 (1회용) */
export async function POST() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    if (!KAKAO_REST_KEY) {
      return NextResponse.json({ error: 'KAKAO_REST_API_KEY 환경변수 미설정' }, { status: 500 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 좌표 없고 주소 있는 출장 건 조회
    const { data: rows, error } = await db
      .from('consultations')
      .select('id, address_road, address_detail')
      .eq('consultation_type', 'field_request')
      .is('latitude', null)
      .not('address_road', 'is', null)
      .limit(50);

    if (error) throw error;
    if (!rows || rows.length === 0) {
      return NextResponse.json({ updated: 0, message: '좌표 채울 건 없음' });
    }

    /** 다중 fallback geocoding:
     *  1) address.json (도로명) → 정확 매칭
     *  2) keyword.json (도로명) → 신축/오래된 주소 유연 매칭
     *  3) keyword.json (도로명 + 상세) → 상호명·건물명 포함된 케이스
     */
    async function geocodeMulti(road: string, detail: string): Promise<{ lat: number; lng: number; via: string } | null> {
      const tryFetch = async (path: string, q: string) => {
        try {
          const res = await fetch(`https://dapi.kakao.com${path}?query=${encodeURIComponent(q)}`, {
            headers: { Authorization: `KakaoAK ${KAKAO_REST_KEY}` },
            signal: AbortSignal.timeout(3000),
          });
          if (!res.ok) return null;
          const json = await res.json();
          const doc = json.documents?.[0];
          if (doc?.y && doc?.x) return { lat: parseFloat(doc.y), lng: parseFloat(doc.x) };
        } catch { /* 다음 시도 */ }
        return null;
      };

      const a = await tryFetch('/v2/local/search/address.json', road);
      if (a) return { ...a, via: 'address' };
      const k = await tryFetch('/v2/local/search/keyword.json', road);
      if (k) return { ...k, via: 'keyword(road)' };
      if (detail) {
        const k2 = await tryFetch('/v2/local/search/keyword.json', `${road} ${detail}`);
        if (k2) return { ...k2, via: 'keyword(road+detail)' };
      }
      return null;
    }

    let updated = 0;
    const results: Array<{ id: string; address: string; lat?: number; lng?: number; via?: string; error?: string }> = [];

    for (const row of rows) {
      const road = (row.address_road || '').trim();
      const detail = (row.address_detail || '').trim();
      if (!road) continue;

      try {
        const geo = await geocodeMulti(road, detail);
        if (geo) {
          await db.from('consultations').update({ latitude: geo.lat, longitude: geo.lng }).eq('id', row.id);
          updated++;
          results.push({ id: row.id, address: road, lat: geo.lat, lng: geo.lng, via: geo.via });
        } else {
          results.push({ id: row.id, address: road, error: 'no result' });
        }

        // Rate limit 방지: 100ms 딜레이
        await new Promise(r => setTimeout(r, 100));
      } catch (err) {
        results.push({ id: row.id, address: road, error: String(err) });
      }
    }

    return NextResponse.json({ updated, total: rows.length, results });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
