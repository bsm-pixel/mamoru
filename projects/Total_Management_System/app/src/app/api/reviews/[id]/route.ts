import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** PATCH /api/reviews/[id] — 리뷰 상태 변경 (인증 필요) */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { status, is_best } = body;

    const updateData: Record<string, unknown> = {};

    // 상태 변경
    if (status) {
      if (!['approved', 'hidden', 'pending'].includes(status)) {
        return NextResponse.json({ error: '유효하지 않은 상태입니다' }, { status: 400 });
      }
      updateData.status = status;
      if (status === 'approved') {
        updateData.approved_at = new Date().toISOString();
      } else {
        updateData.approved_at = null;
      }
    }

    // 베스트 리뷰 토글
    if (typeof is_best === 'boolean') {
      updateData.is_best = is_best;
    }

    if (Object.keys(updateData).length === 0) {
      return NextResponse.json({ error: '변경할 항목이 없습니다' }, { status: 400 });
    }

    const { data, error } = await db
      .from('reviews')
      .update(updateData)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json(data);
  } catch (err) {
    console.error('[reviews] 상태 변경 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
