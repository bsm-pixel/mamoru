import { NextRequest, NextResponse } from 'next/server';
import { requireUser } from '@/lib/auth/require-user';
import { createServiceClient } from '@/lib/supabase/server';
import { isMissingNoticeColumn } from '@/lib/event/campaign-notify';

/** GET /api/campaigns — 캠페인 목록 (CORS도 허용: 고객 폼 projects/event/page_form.html 에서 캠페인명 표시용 — 로그인 없이 공개 유지). POST 는 로그인 필요 */
const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' };
export function OPTIONS() { return new NextResponse(null, { status: 204, headers: CORS }); }

export async function GET() {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (db as any)
      .from('event_campaigns')
      .select('*')
      .order('created_at', { ascending: false });
    if (error) throw error;
    return NextResponse.json({ ok: true, campaigns: data || [] }, { headers: CORS });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS });
  }
}

/** POST /api/campaigns — 캠페인 생성 */
export async function POST(req: NextRequest) {
  const deny = await requireUser(); if (deny) return deny; // 2026-10-29 로그인 없이 열려 있던 관리자 API 차단
  try {
    const body = await req.json();
    if (!body.name?.trim()) return NextResponse.json({ ok: false, error: '캠페인명 필수' }, { status: 400 });
    const db = createServiceClient();
    const row: Record<string, unknown> = {
      name: body.name.trim(),
      type: body.type || 'other',
      status: 'active',
      starts_at: body.starts_at || null,
      ends_at: body.ends_at || null,
      memo: body.memo || null,
      discount_rules: Array.isArray(body.discount_rules) ? body.discount_rules : [],
    };
    // 145: 알림톡 고객 안내 문구 (입력했을 때만 — 마이그 전 컬럼 없음 대비)
    if (typeof body.customer_notice === 'string' && body.customer_notice.trim()) row.customer_notice = body.customer_notice.trim();
    // 152: 무료 이벤트·신청항목 표기
    if (body.payment_type === 'free' || body.payment_type === 'paid') row.payment_type = body.payment_type;
    if (typeof body.items_label === 'string' && body.items_label.trim()) row.items_label = body.items_label.trim();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (db as any).from('event_campaigns').insert(row).select().single();
    if (error) {
      if ('customer_notice' in row && isMissingNoticeColumn(error)) {
        return NextResponse.json({ ok: false, error: '고객 안내 문구 칸이 아직 DB에 없습니다 — 마이그레이션 145 실행 후 다시 저장하세요' }, { status: 409 });
      }
      throw error;
    }
    return NextResponse.json({ ok: true, campaign: data });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
