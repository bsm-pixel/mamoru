import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/serials/audit?serial_id=<uuid>
 *
 * 시리얼 이동 이력 조회 — DB 트리거가 자동으로 product_serial_audit_log 에 기록한 변경 이력.
 * 변경자(changed_by) → profiles JOIN 으로 이름 표시.
 * append-only (DB RLS 로 INSERT/UPDATE/DELETE 거부됨).
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const serialId = req.nextUrl.searchParams.get('serial_id')?.trim();
    if (!serialId) {
      return NextResponse.json({ error: 'serial_id 가 필요합니다' }, { status: 400 });
    }

    // 이력 조회 (최신순, 최근 100건 제한)
    const { data: logs, error } = await db
      .from('product_serial_audit_log')
      .select('*')
      .eq('serial_id', serialId)
      .order('changed_at', { ascending: false })
      .limit(100);

    if (error) {
      return NextResponse.json({ error: error.message }, { status: 500 });
    }

    // 변경자 이름 매핑 (profiles 한 번에 IN 조회 → N+1 방지)
    const userIds = Array.from(
      new Set((logs || []).map((l: { changed_by: string | null }) => l.changed_by).filter(Boolean) as string[])
    );

    let userMap: Record<string, { name: string | null; email: string | null }> = {};
    if (userIds.length > 0) {
      const { data: profiles } = await db
        .from('profiles')
        .select('id, name, email')
        .in('id', userIds);
      userMap = Object.fromEntries(
        (profiles || []).map((p: { id: string; name: string | null; email: string | null }) => [
          p.id,
          { name: p.name, email: p.email },
        ])
      );
    }

    // 응답 가공 — 변경자 이름 포함
    const enriched = (logs || []).map((l: Record<string, unknown>) => ({
      ...l,
      changed_by_name: l.changed_by ? (userMap[l.changed_by as string]?.name || userMap[l.changed_by as string]?.email || '시스템') : '시스템',
    }));

    return NextResponse.json({ logs: enriched });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
