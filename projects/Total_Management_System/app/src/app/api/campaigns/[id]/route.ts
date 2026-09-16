import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { isMissingNoticeColumn } from '@/lib/event/campaign-notify';

/** PATCH /api/campaigns/[id] — 캠페인 수정 (이름/유형/상태/할인규칙) */
export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const body = await req.json();
    const patch: Record<string, unknown> = {};
    if (typeof body.name === 'string') patch.name = body.name.trim();
    if (typeof body.type === 'string') patch.type = body.type;
    if (typeof body.status === 'string') patch.status = body.status;
    if (typeof body.memo === 'string') patch.memo = body.memo;
    // 145: 알림톡 고객 안내 문구 (빈 문자열 = 기본 문구로 되돌림)
    if (typeof body.customer_notice === 'string') patch.customer_notice = body.customer_notice.trim() || null;
    // 152: 무료 이벤트·신청항목 표기
    if (body.payment_type === 'free' || body.payment_type === 'paid') patch.payment_type = body.payment_type;
    if (typeof body.items_label === 'string') patch.items_label = body.items_label.trim() || null;
    if (Array.isArray(body.discount_rules)) patch.discount_rules = body.discount_rules;
    if (body.starts_at !== undefined) patch.starts_at = body.starts_at || null;
    if (body.ends_at !== undefined) patch.ends_at = body.ends_at || null;
    if (Object.keys(patch).length === 0) return NextResponse.json({ ok: true });

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { error } = await (db as any).from('event_campaigns').update(patch).eq('id', id);
    if (error) {
      if ('customer_notice' in patch && isMissingNoticeColumn(error)) {
        return NextResponse.json({ ok: false, error: '고객 안내 문구 칸이 아직 DB에 없습니다 — 마이그레이션 145 실행 후 다시 저장하세요' }, { status: 409 });
      }
      throw error;
    }
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
