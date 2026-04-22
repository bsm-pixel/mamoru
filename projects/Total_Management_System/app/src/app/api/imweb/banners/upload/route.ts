/**
 * POST /api/imweb/banners/upload
 * 배너 이미지 업로드 → Supabase Storage (imweb-banners 버킷)
 *
 * 요청: multipart/form-data
 *   - file: 이미지 파일 (jpg/png/webp, 최대 2MB)
 *   - bannerId: 'main_modal' 등
 *   - oldPath: 이전 이미지 경로 (있으면 삭제)
 *
 * 응답: { ok: true, url, path }
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

const MAX_SIZE = 2 * 1024 * 1024; // 2MB
const ALLOWED_MIME = ['image/jpeg', 'image/png', 'image/webp'];
const BUCKET = 'imweb-banners';

export async function POST(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const form = await req.formData();
    const file = form.get('file') as File | null;
    const bannerId = (form.get('bannerId') as string) || 'main_modal';
    const oldPath = (form.get('oldPath') as string) || '';

    if (!file) {
      return NextResponse.json({ ok: false, error: '파일이 없습니다' }, { status: 400 });
    }

    // 검증
    if (!ALLOWED_MIME.includes(file.type)) {
      return NextResponse.json(
        { ok: false, error: `jpg/png/webp만 업로드 가능 (현재: ${file.type})` },
        { status: 400 },
      );
    }
    if (file.size > MAX_SIZE) {
      return NextResponse.json(
        { ok: false, error: `이미지 크기가 2MB를 초과합니다 (현재: ${(file.size / 1024 / 1024).toFixed(2)}MB)` },
        { status: 400 },
      );
    }

    // 파일명: {bannerId}/{timestamp}-{random}.{ext}
    const ext = file.name.split('.').pop()?.toLowerCase() || 'png';
    const safeExt = ['jpg', 'jpeg', 'png', 'webp'].includes(ext) ? ext : 'png';
    const filename = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${safeExt}`;
    const path = `${bannerId}/${filename}`;

    // Service client로 업로드 (Storage RLS 우회 — 관리자 작업)
    const svc = createServiceClient();

    const { error: uploadErr } = await svc.storage.from(BUCKET).upload(path, file, {
      contentType: file.type,
      upsert: false,
    });

    if (uploadErr) {
      return NextResponse.json(
        { ok: false, error: `업로드 실패: ${uploadErr.message}` },
        { status: 500 },
      );
    }

    // 이전 이미지 삭제 (best-effort)
    if (oldPath && oldPath !== path) {
      try {
        await svc.storage.from(BUCKET).remove([oldPath]);
      } catch {
        /* 이전 이미지 삭제 실패는 무시 */
      }
    }

    // Public URL 생성
    const { data: urlData } = svc.storage.from(BUCKET).getPublicUrl(path);

    return NextResponse.json({
      ok: true,
      url: urlData.publicUrl,
      path,
    });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
