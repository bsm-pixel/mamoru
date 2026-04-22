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
    if (image_url !== undefined) updateData.image_url = image_url;
    if (image_path !== undefined) updateData.image_path = image_path;
    if (link_url !== undefined) updateData.link_url = link_url;
    if (starts_at !== undefined) updateData.starts_at = starts_at;
    if (ends_at !== undefined) updateData.ends_at = ends_at;
    if (dismiss_cookie_hours !== undefined) updateData.dismiss_cookie_hours = Number(dismiss_cookie_hours);

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
