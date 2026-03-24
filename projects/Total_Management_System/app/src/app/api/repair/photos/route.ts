import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

const BUCKET = 'repair-photos';
const MAX_SIZE = 10 * 1024 * 1024; // 10MB

/** GET /api/repair/photos?repair_id=xxx — 사진 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const repairId = req.nextUrl.searchParams.get('repair_id');
    if (!repairId) return NextResponse.json({ error: 'repair_id 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data, error } = await db
      .from('repair_photos')
      .select('*')
      .eq('repair_id', repairId)
      .order('created_at', { ascending: true });

    if (error) throw error;

    // 각 사진에 공개 URL 추가
    const photos = (data || []).map((p: Record<string, unknown>) => ({
      ...p,
      url: supabase.storage.from(BUCKET).getPublicUrl(p.file_path as string).data.publicUrl,
    }));

    return NextResponse.json({ photos });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/repair/photos — 사진 업로드 (multipart/form-data) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const formData = await req.formData();
    const repairId = formData.get('repair_id') as string;
    const file = formData.get('file') as File;
    const memo = formData.get('memo') as string | null;

    if (!repairId || !file) {
      return NextResponse.json({ error: 'repair_id와 file 필수' }, { status: 400 });
    }
    if (file.size > MAX_SIZE) {
      return NextResponse.json({ error: '파일 크기 10MB 초과' }, { status: 400 });
    }

    // Storage에 업로드
    const ext = file.name.split('.').pop() || 'jpg';
    const fileName = `${Date.now()}.${ext}`;
    const filePath = `repairs/${repairId}/${fileName}`;

    const { error: uploadErr } = await supabase.storage
      .from(BUCKET)
      .upload(filePath, file, { contentType: file.type, upsert: false });

    if (uploadErr) throw uploadErr;

    // DB에 메타 저장
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data, error } = await db
      .from('repair_photos')
      .insert({
        repair_id: repairId,
        file_path: filePath,
        file_name: file.name,
        file_size: file.size,
        memo: memo || null,
        uploaded_by: user.id,
      })
      .select()
      .single();

    if (error) throw error;

    const url = supabase.storage.from(BUCKET).getPublicUrl(filePath).data.publicUrl;

    return NextResponse.json({ photo: { ...data, url } });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/repair/photos?id=xxx — 사진 삭제 */
export async function DELETE(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const photoId = req.nextUrl.searchParams.get('id');
    if (!photoId) return NextResponse.json({ error: 'id 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // DB에서 경로 조회
    const { data: photo } = await db
      .from('repair_photos')
      .select('file_path')
      .eq('id', photoId)
      .single();

    if (photo?.file_path) {
      await supabase.storage.from(BUCKET).remove([photo.file_path]);
    }

    await db.from('repair_photos').delete().eq('id', photoId);

    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
