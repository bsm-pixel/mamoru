/**
 * POST /api/imweb/script-install
 * 아임웹 Script API로 우리 widget.js 태그를 자동 등록/업데이트
 *
 * Body: { unitCode: string, position?: 'header'|'body'|'footer' }
 *       position 기본 'footer' (defer 효과 + 페이지 로딩 성능)
 *
 * 동작:
 *   1. OAuth 토큰 + script:write 스코프 확인
 *   2. GET /script로 기존 등록 여부 확인
 *   3. 있으면 PUT, 없으면 POST
 *   4. system_settings에 unit_code/position 기록
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { upsertMamoruWidget, type ScriptPosition } from '@/lib/imweb/script-api';

const BASE_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://app-eta-sandy-75.vercel.app';
const WIDGET_URL = `${BASE_URL}/api/imweb/banner-widget.js`;

export async function POST(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const body = await req.json();
    const unitCode = String(body.unitCode || '').trim();
    const position = (body.position || 'footer') as ScriptPosition;

    if (!unitCode) {
      return NextResponse.json({ ok: false, error: 'unitCode가 필요합니다' }, { status: 400 });
    }
    if (!['header', 'body', 'footer'].includes(position)) {
      return NextResponse.json({ ok: false, error: 'position은 header/body/footer 중 하나' }, { status: 400 });
    }

    const result = await upsertMamoruWidget(unitCode, position, WIDGET_URL);

    // 설정 저장 (다음번 자동 갱신용)
    const svc = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const svcAny = svc as any;
    const now = new Date().toISOString();
    await svcAny.from('system_settings').upsert(
      { key: 'imweb.banner_widget.unit_code', value: unitCode, updated_at: now },
      { onConflict: 'key' },
    );
    await svcAny.from('system_settings').upsert(
      { key: 'imweb.banner_widget.position', value: position, updated_at: now },
      { onConflict: 'key' },
    );
    await svcAny.from('system_settings').upsert(
      { key: 'imweb.banner_widget.installed_at', value: now, updated_at: now },
      { onConflict: 'key' },
    );

    return NextResponse.json({
      ok: true,
      data: {
        action: result.action, // 'created' | 'updated'
        unitCode,
        position,
        widgetUrl: WIDGET_URL,
      },
    });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    // scope 부족 에러 감지
    if (/30099|30103|scope|토큰|refresh/i.test(msg)) {
      return NextResponse.json(
        {
          ok: false,
          error: msg,
          needsReauth: true,
          hint: '아임웹 OAuth를 script:write 스코프 포함하여 재연결이 필요합니다.',
        },
        { status: 403 },
      );
    }
    return NextResponse.json({ ok: false, error: msg }, { status: 500 });
  }
}

/** GET /api/imweb/script-install — 현재 설치 상태 조회 */
export async function GET() {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const svc = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const svcAny = svc as any;
  const { data } = await svcAny
    .from('system_settings')
    .select('key, value')
    .in('key', [
      'imweb.banner_widget.unit_code',
      'imweb.banner_widget.position',
      'imweb.banner_widget.installed_at',
    ]);

  const map: Record<string, string> = {};
  for (const r of data || []) if (r.value) map[r.key] = String(r.value).replace(/^"|"$/g, '');

  return NextResponse.json({
    ok: true,
    data: {
      unit_code: map['imweb.banner_widget.unit_code'] || '',
      position: map['imweb.banner_widget.position'] || 'footer',
      installed_at: map['imweb.banner_widget.installed_at'] || '',
    },
  });
}
