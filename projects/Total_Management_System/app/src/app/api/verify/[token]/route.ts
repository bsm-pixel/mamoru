import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/verify/[token] — 정품확인 공개 API (인증 불필요)
 *
 * verify_token으로 시리얼 조회 → 제품 정보 반환
 * 시리얼번호는 반환하지 않음 (보안)
 */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ token: string }> },
) {
  try {
    const { token } = await params;

    if (!token || token.length < 8) {
      return NextResponse.json({ valid: false, error: '유효하지 않은 토큰' }, { status: 400 });
    }

    const supabase = await createServerSupabaseClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // verify_token으로 시리얼 + 제품 조회
    const { data: serial, error } = await db
      .from('product_serials')
      .select(`
        id,
        status,
        sold_via,
        sold_to_name,
        sold_at,
        manufactured_at,
        created_at,
        product_id,
        products:product_id (
          name,
          sku,
          category,
          description,
          image_url
        )
      `)
      .eq('verify_token', token)
      .single();

    if (error || !serial) {
      return NextResponse.json({
        valid: false,
        message: '등록되지 않은 제품입니다. MAMORU 정품이 아닐 수 있습니다.',
      });
    }

    const product = serial.products;

    // 판매 채널 한글화
    const channelLabel: Record<string, string> = {
      offline: 'MAMORU 공식 매장',
      online: 'MAMORU 공식 온라인몰',
      contract: 'MAMORU 공식 계약',
    };

    // 딜러 판매 여부 확인
    let soldViaLabel = channelLabel[serial.sold_via] || null;
    if (serial.sold_via === 'offline' && serial.sold_to_name) {
      // 딜러 판매인 경우 고객 타입 확인
      const { data: customer } = await db
        .from('customers')
        .select('customer_type')
        .eq('name', serial.sold_to_name)
        .limit(1)
        .single();

      if (customer?.customer_type === 'dealer' || customer?.customer_type === 'academy') {
        soldViaLabel = `MAMORU 공식 딜러 (${serial.sold_to_name})`;
      }
    }

    return NextResponse.json({
      valid: true,
      message: 'MAMORU 정품 인증 확인',
      product: {
        name: product?.name || '알 수 없는 제품',
        sku: product?.sku || null,
        category: product?.category || null,
        description: product?.description || null,
        image_url: product?.image_url || null,
      },
      details: {
        status: serial.status === 'sold' ? '판매완료' : '미판매',
        sold_via: soldViaLabel,
        sold_at: serial.sold_at || null,
        manufactured_at: serial.manufactured_at || null,
        registered_at: serial.created_at,
      },
    });
  } catch (err) {
    return NextResponse.json({ valid: false, error: String(err) }, { status: 500 });
  }
}
