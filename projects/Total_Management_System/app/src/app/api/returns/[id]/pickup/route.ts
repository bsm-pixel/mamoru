import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { bookReturnPickupFor } from '@/lib/returns/book-pickup';

/** POST /api/returns/[id]/pickup — 롯데 반품 수거접수 (ustRtgSctCd='02', 고객집 → 마모루 회수)
 *  로직 SSOT = lib/returns/book-pickup.ts (납품 반품 접수와 공용, 2026-10-21 분리 — 동작 동일)
 *  ⚠️ 접수 후 취소는 ALPS 화면에서 수동(취소 API 없음).
 */
export async function POST(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const r = await bookReturnPickupFor(db, id);
    if (!r.ok) return NextResponse.json({ error: r.error }, { status: r.status });
    if (r.warning) return NextResponse.json({ success: true, warning: r.warning, invoiceNumber: r.invoiceNumber });
    return NextResponse.json({ success: true, return: r.row });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
