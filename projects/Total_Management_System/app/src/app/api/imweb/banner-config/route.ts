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
      .select('id, enabled, title, description, image_url, link_url, images, starts_at, ends_at, dismiss_cookie_hours, updated_at')
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

    // images 배열 정규화 (Phase 2: 슬라이드 지원)
    // DB에 images가 비어있으면 legacy image_url을 1개짜리 배열로 fallback
    let imagesArr: Array<{ url: string; link_url?: string }> = [];
    if (Array.isArray(data.images) && data.images.length > 0) {
      imagesArr = data.images
        .filter((img: { url?: string }) => img && img.url)
        .map((img: { url: string; link_url?: string }) => ({
          url: img.url,
          link_url: img.link_url || '',
        }));
    } else if (data.image_url) {
      imagesArr = [{ url: data.image_url, link_url: data.link_url || '' }];
    }

    // 이미지 하나도 없으면 배너 자체를 미노출 (텍스트 전용 배너는 MVP 불가)
    if (imagesArr.length === 0) {
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
        images: imagesArr,
        // backwards-compat: 구 widget도 동작하도록 첫 이미지를 레거시 필드로 노출
        image_url: imagesArr[0].url,
        link_url: imagesArr[0].link_url || '',
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
