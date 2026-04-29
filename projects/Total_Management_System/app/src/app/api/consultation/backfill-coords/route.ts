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

    let updated = 0;
    const results: Array<{ id: string; address: string; lat?: number; lng?: number; error?: string }> = [];

    for (const row of rows) {
      // 카카오 주소검색 API는 도로명/지번만 인식. address_detail은 매칭 실패 원인이라 제외.
      const address = (row.address_road || '').trim();
      if (!address) continue;

      try {
        const url = `https://dapi.kakao.com/v2/local/search/address.json?query=${encodeURIComponent(address)}`;
        const res = await fetch(url, {
          headers: { Authorization: `KakaoAK ${KAKAO_REST_KEY}` },
          signal: AbortSignal.timeout(3000),
        });

        if (!res.ok) {
          results.push({ id: row.id, address, error: `HTTP ${res.status}` });
          continue;
        }

        const json = await res.json();
        const doc = json.documents?.[0];
        if (doc?.y && doc?.x) {
          const lat = parseFloat(doc.y);
          const lng = parseFloat(doc.x);
          await db.from('consultations').update({ latitude: lat, longitude: lng }).eq('id', row.id);
          updated++;
          results.push({ id: row.id, address, lat, lng });
        } else {
          results.push({ id: row.id, address, error: 'no result' });
        }

        // Rate limit 방지: 100ms 딜레이
        await new Promise(r => setTimeout(r, 100));
      } catch (err) {
        results.push({ id: row.id, address, error: String(err) });
      }
    }

    return NextResponse.json({ updated, total: rows.length, results });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
