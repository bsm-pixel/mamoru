import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/repair/public/holidays?years=2025,2026 — 공휴일 조회 (비인증) */
export async function GET(req: NextRequest) {
  try {
    const yearsParam = req.nextUrl.searchParams.get('years');
    const years = yearsParam
      ? yearsParam.split(',').map(y => parseInt(y.trim())).filter(y => !isNaN(y))
      : [new Date().getFullYear()];

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data, error } = await dbAny
      .from('holidays')
      .select('date, name, year')
      .in('year', years)
      .order('date', { ascending: true });

    if (error) throw error;

    return NextResponse.json(
      { ok: true, data: data || [] },
      { headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[repair/public/holidays] 조회 실패:', err);
    return NextResponse.json(
      { ok: false, error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
