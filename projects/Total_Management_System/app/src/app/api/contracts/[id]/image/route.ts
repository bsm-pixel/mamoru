import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/contracts/[id]/image — 계약서 이미지 업로드 */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { imageBase64 } = body as { imageBase64: string };

    if (!imageBase64) {
      return NextResponse.json({ error: 'imageBase64 required' }, { status: 400 });
    }

    // base64 → Buffer
    const base64Data = imageBase64.replace(/^data:image\/\w+;base64,/, '');
    const buffer = Buffer.from(base64Data, 'base64');

    const fileName = `contracts/${id}.png`;

    // Supabase Storage 업로드 (contracts 버킷)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { error: uploadError } = await db.storage
      .from('contracts')
      .upload(fileName, buffer, {
        contentType: 'image/png',
        upsert: true,
      });

    if (uploadError) throw uploadError;

    // Public URL
    const { data: urlData } = db.storage.from('contracts').getPublicUrl(fileName);
    const imageUrl = urlData?.publicUrl || '';

    // contracts.image_url 업데이트
    const { error: updateError } = await db
      .from('contracts')
      .update({ image_url: imageUrl })
      .eq('id', id);

    if (updateError) throw updateError;

    return NextResponse.json({ imageUrl });
  } catch (err) {
    console.error('[contracts/image POST] error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
