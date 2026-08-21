import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { randomBytes } from 'crypto';

/**
 * POST /api/serials/[id]/verify-label — 정품 QR 라벨용 데이터
 *   시리얼의 verify_token 을 보장(없으면 생성)하고, QR URL + 모델명(SKU) 반환.
 *   QR URL = page.mamoru.kr/projects/verify/?t=<verify_token>  (verify 페이지가 ?t 읽음)
 */
export async function POST(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;

  const auth = await createServerSupabaseClient();
  const { data: { user } } = await auth.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db: any = createServiceClient();
  const { data: serial, error } = await db
    .from('product_serials')
    .select('id, serial_number, verify_token, product_id')
    .eq('id', id)
    .single();
  if (error || !serial) return NextResponse.json({ error: '시리얼을 찾을 수 없습니다' }, { status: 404 });

  let token: string = serial.verify_token;
  if (!token) {
    token = randomBytes(12).toString('hex'); // 24 hex — 추측 불가
    const { error: updErr } = await db.from('product_serials').update({ verify_token: token }).eq('id', id);
    if (updErr) return NextResponse.json({ error: `토큰 저장 실패: ${updErr.message || updErr}` }, { status: 500 });
  }

  let model = '';
  if (serial.product_id) {
    const { data: prod } = await db.from('products').select('sku, name').eq('id', serial.product_id).single();
    model = (prod?.sku || prod?.name || '').toString();
  }

  return NextResponse.json({
    verifyUrl: `https://page.mamoru.kr/projects/verify/?t=${token}`,
    model,
    serial: serial.serial_number,
  });
}
