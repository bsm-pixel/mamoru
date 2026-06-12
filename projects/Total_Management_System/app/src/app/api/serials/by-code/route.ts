import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { parseSerial, formatSerial } from '@/lib/serial/format';

/**
 * GET /api/serials/by-code?code=시리얼 — 스캔용 경량 조회.
 * 하이픈/대소문자 무시. 반환: { found, serial:{id,serial_number,product_id,status,offline_sale_id,sale_item_id,sold_to_name} }
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const code = req.nextUrl.searchParams.get('code')?.trim();
    if (!code) return NextResponse.json({ found: false });

    // 하이픈 유무 무시 — 입력값 + M{YY}-…정식형 둘 다
    const candidates = [code];
    const p = parseSerial(code);
    if (p) { const c = formatSerial(p.year2, p.seq); if (c !== code) candidates.push(c); }

    const { data } = await db
      .from('product_serials')
      .select('id, serial_number, product_id, status, offline_sale_id, sale_item_id, sold_to_name')
      .in('serial_number', candidates)
      .limit(1);

    const serial = data?.[0];
    if (!serial) return NextResponse.json({ found: false });
    return NextResponse.json({ found: true, serial });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}
