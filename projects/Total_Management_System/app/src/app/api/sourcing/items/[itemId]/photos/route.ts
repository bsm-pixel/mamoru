import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

const BUCKET = 'repair-photos'; // 기존 공개 버킷 재사용 (sourcing/ 경로 분리)
const MAX_SIZE = 10 * 1024 * 1024;

/** POST /api/sourcing/items/[itemId]/photos — 입고 사진 업로드(multipart) → inbound_photos 에 URL append */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ itemId: string }> },
) {
  try {
    const { itemId } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const formData = await req.formData();
    const file = formData.get('file') as File | null;
    if (!file) return NextResponse.json({ error: 'file 필수' }, { status: 400 });
    if (file.size > MAX_SIZE) return NextResponse.json({ error: '파일 10MB 초과' }, { status: 400 });

    const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
    const filePath = `sourcing/${itemId}/${Date.now()}.${ext}`;

    const { error: upErr } = await supabase.storage
      .from(BUCKET)
      .upload(filePath, file, { contentType: file.type, upsert: false });
    if (upErr) throw upErr;

    const url = supabase.storage.from(BUCKET).getPublicUrl(filePath).data.publicUrl;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: cur, error: getErr } = await db
      .from('sourcing_items')
      .select('inbound_photos')
      .eq('id', itemId)
      .single();
    if (getErr) throw getErr;

    const photos: string[] = Array.isArray(cur.inbound_photos) ? cur.inbound_photos : [];
    const next = [...photos, url].slice(0, 5);

    const { error: updErr } = await db
      .from('sourcing_items')
      .update({ inbound_photos: next, updated_at: new Date().toISOString() })
      .eq('id', itemId);
    if (updErr) throw updErr;

    return NextResponse.json({ url, inbound_photos: next }, { status: 201 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/sourcing/items/[itemId]/photos?url=... — inbound_photos 에서 1장 제거 */
export async function DELETE(
  req: NextRequest,
  { params }: { params: Promise<{ itemId: string }> },
) {
  try {
    const { itemId } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const target = req.nextUrl.searchParams.get('url');
    if (!target) return NextResponse.json({ error: 'url 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: cur, error: getErr } = await db
      .from('sourcing_items')
      .select('inbound_photos')
      .eq('id', itemId)
      .single();
    if (getErr) throw getErr;

    const photos: string[] = Array.isArray(cur.inbound_photos) ? cur.inbound_photos : [];
    const next = photos.filter((p) => p !== target);

    const { error: updErr } = await db
      .from('sourcing_items')
      .update({ inbound_photos: next, updated_at: new Date().toISOString() })
      .eq('id', itemId);
    if (updErr) throw updErr;

    return NextResponse.json({ inbound_photos: next });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
