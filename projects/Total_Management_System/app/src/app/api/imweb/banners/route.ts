/**
 * /api/imweb/banners
 *   GET    — 배너 전체 조회 (관리자)
 *   PATCH  — 배너 업데이트 (관리자)
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

export async function GET() {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = supabase as any;
  const { data, error } = await db
    .from('imweb_banners')
    .select('*')
    .order('id', { ascending: true });

  if (error) return NextResponse.json({ ok: false, error: error.message }, { status: 500 });
  return NextResponse.json({ ok: true, data });
}

const MAX_IMAGES = 5;

interface BannerImage {
  url: string;
  path?: string;
  link_url?: string;
}

/** images 배열 검증 + 정규화 */
function validateImages(raw: unknown): { ok: true; images: BannerImage[] } | { ok: false; error: string } {
  if (!Array.isArray(raw)) return { ok: false, error: 'images는 배열이어야 합니다' };
  if (raw.length > MAX_IMAGES) return { ok: false, error: `이미지는 최대 ${MAX_IMAGES}장까지 업로드 가능합니다` };

  const normalized: BannerImage[] = [];
  for (let i = 0; i < raw.length; i++) {
    const item = raw[i] as Record<string, unknown>;
    if (!item || typeof item !== 'object') {
      return { ok: false, error: `images[${i}]가 객체가 아닙니다` };
    }
    const url = typeof item.url === 'string' ? item.url : '';
    if (!url) return { ok: false, error: `images[${i}].url이 필요합니다` };
    normalized.push({
      url,
      path: typeof item.path === 'string' ? item.path : '',
      link_url: typeof item.link_url === 'string' ? item.link_url : '',
    });
  }
  return { ok: true, images: normalized };
}

export async function PATCH(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const body = await req.json();
    const {
      id,
      enabled,
      title,
      description,
      image_url,
      image_path,
      link_url,
      images,
      starts_at,
      ends_at,
      dismiss_cookie_hours,
    } = body;

    if (!id) {
      return NextResponse.json({ ok: false, error: 'id가 필요합니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const updateData: Record<string, unknown> = {
      updated_at: new Date().toISOString(),
      updated_by: user.id,
    };

    // 명시적으로 전달된 필드만 업데이트 (undefined는 무시, null은 반영)
    if (enabled !== undefined) updateData.enabled = !!enabled;
    if (title !== undefined) updateData.title = title;
    if (description !== undefined) updateData.description = description;
    if (starts_at !== undefined) updateData.starts_at = starts_at;
    if (ends_at !== undefined) updateData.ends_at = ends_at;
    if (dismiss_cookie_hours !== undefined) updateData.dismiss_cookie_hours = Number(dismiss_cookie_hours);

    // images 배열 처리 (Phase 2: 슬라이드 지원)
    if (images !== undefined) {
      const validation = validateImages(images);
      if (!validation.ok) {
        return NextResponse.json({ ok: false, error: validation.error }, { status: 400 });
      }
      updateData.images = validation.images;

      // backwards-compat: 첫 번째 이미지를 legacy 컬럼에 동기화
      const first = validation.images[0];
      updateData.image_url = first?.url || null;
      updateData.image_path = first?.path || null;
      updateData.link_url = first?.link_url || null;
    } else {
      // images 미전달 시에만 legacy 필드 개별 처리 (단일 이미지 모드)
      if (image_url !== undefined) updateData.image_url = image_url;
      if (image_path !== undefined) updateData.image_path = image_path;
      if (link_url !== undefined) updateData.link_url = link_url;
    }

    const { data, error } = await db
      .from('imweb_banners')
      .update(updateData)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json({ ok: true, data });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
