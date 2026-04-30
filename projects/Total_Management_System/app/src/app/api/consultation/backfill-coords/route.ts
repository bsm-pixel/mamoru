import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { geocodeForConsultation } from '@/lib/kakao/geocode';

/** POST /api/consultation/backfill-coords — 좌표 없는 출장 건에 좌표 채우기
 *  - 회귀 방지: lib/kakao/geocode 공통 helper 사용 (3단계 fallback)
 *  - field-request-map 마운트 시 자동 호출됨 → 사장님이 지도 열기만 하면 NULL 자동 복구
 */
export async function POST() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 좌표 없고 주소 있는 출장 건 조회 (cancelled 제외)
    const { data: rows, error } = await db
      .from('consultations')
      .select('id, address_road, address_detail')
      .eq('consultation_type', 'field_request')
      .is('latitude', null)
      .not('address_road', 'is', null)
      .neq('status', 'cancelled')
      .limit(50);

    if (error) throw error;
    if (!rows || rows.length === 0) {
      return NextResponse.json({ updated: 0, message: '좌표 채울 건 없음' });
    }

    let updated = 0;
    const results: Array<{ id: string; address: string; lat?: number; lng?: number; via?: string; error?: string }> = [];

    for (const row of rows) {
      const road = (row.address_road || '').trim();
      const detail = (row.address_detail || '').trim();
      if (!road) continue;

      try {
        const geo = await geocodeForConsultation(road, detail);
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
