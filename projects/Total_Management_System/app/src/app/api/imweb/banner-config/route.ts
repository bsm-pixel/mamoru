/**
 * GET /api/imweb/banner-config
 * 공개 API (CORS) — 아임웹 사이트의 widget.js가 호출
 *
 * 응답: 활성 배너의 공개 정보만 반환
 *   enabled: false → 즉시 반환 (배너 미표시)
 *   기간(starts_at/ends_at) 외면 → enabled: false로 간주
 *
 * 캐시: 60초 (사장님 변경 후 1분 내 반영)
 */

import { NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

const CACHE_HEADERS = {
  'Cache-Control': 'public, max-age=60, s-maxage=60',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

export async function GET() {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data, error } = await dbAny
      .from('imweb_banners')
      .select('id, enabled, title, description, image_url, link_url, starts_at, ends_at, dismiss_cookie_hours, updated_at')
      .eq('id', 'main_modal')
      .single();

    if (error || !data) {
      return NextResponse.json(
        { enabled: false },
        { headers: { ...CORS_HEADERS, ...CACHE_HEADERS } },
      );
    }

    const now = new Date();
    const startOk = !data.starts_at || new Date(data.starts_at) <= now;
    const endOk = !data.ends_at || new Date(data.ends_at) >= now;
    const active = !!data.enabled && startOk && endOk;

    if (!active) {
      return NextResponse.json(
        { enabled: false },
        { headers: { ...CORS_HEADERS, ...CACHE_HEADERS } },
      );
    }

    return NextResponse.json(
      {
        enabled: true,
        title: data.title || '',
        description: data.description || '',
        image_url: data.image_url || '',
        link_url: data.link_url || '',
        dismiss_cookie_hours: data.dismiss_cookie_hours || 24,
        version: data.updated_at || new Date().toISOString(),
      },
      { headers: { ...CORS_HEADERS, ...CACHE_HEADERS } },
    );
  } catch {
    return NextResponse.json(
      { enabled: false },
      { headers: { ...CORS_HEADERS, ...CACHE_HEADERS } },
    );
  }
}
