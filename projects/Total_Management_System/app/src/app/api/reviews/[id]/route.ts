import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isReviewType, isAllowedReviewSubtype } from '@/lib/reviews/consult-subtype';

/** 수정 가능한 필드 목록 */
const EDITABLE_FIELDS = ['name', 'stars', 'content', 'photo_urls', 'type', 'subtype', 'product', 'created_at'] as const;

/** PATCH /api/reviews/[id] — 리뷰 수정 (상태/베스트/내용/사진 등) */
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
    const { status, is_best, ...rest } = body;

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

    // 편집 가능 필드 반영
    for (const field of EDITABLE_FIELDS) {
      if (rest[field] !== undefined) {
        updateData[field] = rest[field];
      }
    }

    // 유형/세부유형 정합성 (2026-10-20) — 흐름을 잘못 골라 보낸 후기를 고칠 때
    //  ① subtype 은 유형별 허용값만 ② 상담/복원수리로 옮기면 제품명·제품군 제거(후기벽 제품 칩·상품 위젯 노출 차단)
    if (rest.type !== undefined || rest.subtype !== undefined) {
      let nextType: unknown = rest.type;
      if (nextType === undefined) {
        const { data: cur } = await db.from('reviews').select('type').eq('id', id).single();
        nextType = cur?.type;
      }
      if (!isReviewType(nextType)) {
        return NextResponse.json({ error: '유효하지 않은 리뷰 유형입니다' }, { status: 400 });
      }
      if (rest.type !== undefined && rest.subtype === undefined) {
        return NextResponse.json({ error: '유형을 바꿀 때는 세부유형도 함께 지정해야 합니다' }, { status: 400 });
      }
      if (!isAllowedReviewSubtype(nextType, rest.subtype)) {
        return NextResponse.json({ error: '이 유형에 맞지 않는 세부유형입니다' }, { status: 400 });
      }
      updateData.subtype = rest.subtype || null;
      if (nextType !== 'purchase') {
        updateData.product = null;
        updateData.product_group = null;
      }
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
    console.error('[reviews] 수정 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/reviews/[id] — 리뷰 삭제 */
export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { error } = await db.from('reviews').delete().eq('id', id);
    if (error) throw error;

    return NextResponse.json({ success: true });
  } catch (err) {
    console.error('[reviews] 삭제 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
