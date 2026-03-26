import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { randomBytes } from 'crypto';

/** verify_token 생성 — 12자리 hex (URL-safe, 추측 불가) */
function generateVerifyToken(): string {
  return randomBytes(6).toString('hex'); // 6바이트 = 12자리 hex
}

/**
 * POST /api/serials/batch — 시리얼 일괄 생성
 *
 * body: { product_id, count, start_number, lot_number? }
 * - start_number: 시작 번호 (이지캐드 연번과 동일하게 입력)
 * - 포맷: 순차 8자리 숫자 (예: 13792241, 13792242, ...)
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { product_id, count, start_number, lot_number } = await req.json() as {
      product_id: string;
      count: number;
      start_number: number;
      lot_number?: string;
    };

    if (!product_id || !count || count < 1 || count > 100) {
      return NextResponse.json({ error: '1~100개 범위로 입력해주세요' }, { status: 400 });
    }

    if (!start_number || start_number < 1) {
      return NextResponse.json({ error: '시작 번호를 입력해주세요' }, { status: 400 });
    }

    // 순차 시리얼 생성
    const serials = Array.from({ length: count }, (_, i) => {
      const serialNumber = String(start_number + i).padStart(8, '0');
      return {
        product_id,
        serial_number: serialNumber,
        barcode: serialNumber,
        verify_token: generateVerifyToken(),
        lot_number: lot_number || null,
        created_by: user.id,
      };
    });

    // DB UNIQUE 제약이 최종 안전장치 (serial_number, barcode, verify_token 모두 UNIQUE)
    const { error } = await db
      .from('product_serials')
      .insert(serials);

    if (error) throw error;

    const startStr = String(start_number).padStart(8, '0');
    const endStr = String(start_number + count - 1).padStart(8, '0');

    return NextResponse.json({
      created: count,
      range: `${startStr} ~ ${endStr}`,
      start_number,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
