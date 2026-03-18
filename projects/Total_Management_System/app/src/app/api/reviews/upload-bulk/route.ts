import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

const MAX_SIZE = 5 * 1024 * 1024; // 5MB per file
const ALLOWED_TYPES = ['image/jpeg', 'image/png', 'image/webp'];

/** POST /api/reviews/upload-bulk — 사진 일괄 업로드 (인증 필요) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const formData = await req.formData();
    const files = formData.getAll('files') as File[];

    if (!files || files.length === 0) {
      return NextResponse.json({ error: '파일이 없습니다' }, { status: 400 });
    }

    const db = createServiceClient();
    const results: { name: string; url?: string; error?: string }[] = [];

    for (const file of files) {
      try {
        if (!ALLOWED_TYPES.includes(file.type)) {
          results.push({ name: file.name, error: '지원하지 않는 형식' });
          continue;
        }
        if (file.size > MAX_SIZE) {
          results.push({ name: file.name, error: '5MB 초과' });
          continue;
        }

        const ext = file.type === 'image/png' ? 'png' : file.type === 'image/webp' ? 'webp' : 'jpg';
        const fileName = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`;
        const filePath = `reviews/${fileName}`;

        const buffer = Buffer.from(await file.arrayBuffer());

        const { error: uploadError } = await db.storage
          .from('review-photos')
          .upload(filePath, buffer, {
            contentType: file.type,
            upsert: false,
          });

        if (uploadError) throw uploadError;

        const { data: urlData } = db.storage
          .from('review-photos')
          .getPublicUrl(filePath);

        results.push({ name: file.name, url: urlData.publicUrl });
      } catch (err) {
        results.push({ name: file.name, error: String(err) });
      }
    }

    const uploaded = results.filter(r => r.url).length;
    return NextResponse.json({ uploaded, total: files.length, results });
  } catch (err) {
    console.error('[reviews/upload-bulk] 업로드 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
