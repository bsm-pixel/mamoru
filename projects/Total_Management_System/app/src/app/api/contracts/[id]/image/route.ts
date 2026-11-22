import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { requireUser } from '@/lib/auth/require-user';
import { errMsg } from '@/lib/utils/err';
import { contractImagePath } from '@/lib/contracts/image-path';

/**
 * 계약서 이미지 (서명 · 고객 정보가 들어 있다)
 *
 * 2026-11-22 공개 주소 → 짧은 수명 주소로 (작업실 t-261121-tms-leftover)
 *   전엔 누구나 열 수 있는 공개 주소(getPublicUrl)를 contracts.image_url 에 저장했다.
 *   이제 저장소 안 위치(contracts/<id>.png)만 저장하고, 열 때 로그인을 확인한 뒤 1분짜리 서명 주소로 보낸다
 *   (경비 영수증 api/expenses/receipt 와 같은 방식). 'contracts' 저장소는 비공개로 바꾼다.
 */
const BUCKET = 'contracts';

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
      .from(BUCKET)
      .upload(fileName, buffer, {
        contentType: 'image/png',
        upsert: true,
      });

    if (uploadError) throw uploadError;

    // 저장소 안 위치만 저장한다(공개 주소 X) — 열 때는 GET 이 서명 주소로 보낸다
    const { error: updateError } = await db
      .from('contracts')
      .update({ image_url: fileName })
      .eq('id', id);

    if (updateError) throw updateError;

    return NextResponse.json({ imageUrl: `/api/contracts/${id}/image` });
  } catch (err) {
    console.error('[contracts/image POST] error:', err);
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}

/** GET /api/contracts/[id]/image — 로그인 확인 뒤 1분짜리 주소로 이동 */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const deny = await requireUser(); if (deny) return deny;
  try {
    const { id } = await params;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { data: c } = await db.from('contracts').select('image_url').eq('id', id).maybeSingle();
    const path = contractImagePath(c?.image_url);
    if (!path) return NextResponse.json({ error: '계약서 이미지가 없습니다' }, { status: 404 });
    const { data, error } = await db.storage.from(BUCKET).createSignedUrl(path, 60);
    if (error || !data?.signedUrl) return NextResponse.json({ error: String(error?.message || '주소 생성 실패') }, { status: 500 });
    return NextResponse.redirect(data.signedUrl);
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
