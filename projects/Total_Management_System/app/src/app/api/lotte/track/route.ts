import { NextRequest, NextResponse } from 'next/server';
import { queryStatus } from '@/lib/lotte/client';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/lotte/track?invNo=xxx — 배송 조회 */
export async function GET(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const invNo = request.nextUrl.searchParams.get('invNo') || '';
    const result = await queryStatus(invNo);
    return NextResponse.json(result);
  } catch (err) {
    console.error('[lotte/track] 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
