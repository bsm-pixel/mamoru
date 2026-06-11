import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { randomBytes } from 'crypto';
import { nextSerialNumber, nextSerialBatch } from '@/lib/serial/next-serial';

/** verify_token 생성 — 12자리 hex */
function generateVerifyToken(): string {
  return randomBytes(6).toString('hex');
}

/**
 * GET /api/serials/batch — 다음 시리얼 번호 조회 (M{YY}-{NNNN})
 */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const next = await nextSerialNumber(db);
    return NextResponse.json({ next_serial: next });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/**
 * POST /api/serials/batch — 시리얼 일괄 생성
 *
 * 보관(raw_stock)에서 차감 → 시리얼 생성 (zone: ready/display)
 * raw_stock 부족 시 거부
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { product_id, count, lot_number, warehouse_zone } = await req.json() as {
      product_id: string;
      count: number;
      lot_number?: string;
      warehouse_zone?: 'raw' | 'ready' | 'display';
    };

    if (!product_id || !count || count < 1 || count > 100) {
      return NextResponse.json({ error: '1~100개 범위로 입력해주세요' }, { status: 400 });
    }

    // 제품 조회 — raw_stock 확인
    const { data: product } = await db
      .from('products')
      .select('id, raw_stock, stock_quantity, imweb_product_no')
      .eq('id', product_id)
      .single();

    if (!product) {
      return NextResponse.json({ error: '제품을 찾을 수 없습니다' }, { status: 404 });
    }

    const rawStock = product.raw_stock || 0;
    if (rawStock < count) {
      return NextResponse.json({
        error: `보관창고 재고가 부족합니다 (보관: ${rawStock}개, 요청: ${count}개)`,
      }, { status: 400 });
    }

    // 시리얼 번호 목록 생성 — M{YY}-{NNNN} 연속 (서버에서 다음 번호부터 자동)
    const serialNumbers = await nextSerialBatch(db, count);

    // 중복 체크
    const { data: existing } = await db
      .from('product_serials')
      .select('serial_number')
      .in('serial_number', serialNumbers);

    if (existing && existing.length > 0) {
      const duplicates = existing.map((e: { serial_number: string }) => e.serial_number);
      return NextResponse.json({
        error: `이미 존재하는 시리얼번호가 ${duplicates.length}개 있습니다: ${duplicates.join(', ')}`,
        duplicates,
      }, { status: 409 });
    }

    // zone 결정 — 시리얼 생성은 ready 또는 display (보관에서 꺼내는 것이므로)
    const validZones = ['ready', 'display'];
    const zone = warehouse_zone && validZones.includes(warehouse_zone) ? warehouse_zone : 'ready';

    // 시리얼 생성
    const serials = serialNumbers.map((serialNumber) => ({
      product_id,
      serial_number: serialNumber,
      barcode: serialNumber,
      verify_token: generateVerifyToken(),
      warehouse_zone: zone,
      lot_number: lot_number || null,
      created_by: user.id,
    }));

    const { error: insertErr } = await db.from('product_serials').insert(serials);
    if (insertErr) throw insertErr;

    // 보관 수량 차감 (raw_stock -= count, stock_quantity는 유지 — 총 재고는 변하지 않음)
    const { error: updateErr } = await db
      .from('products')
      .update({
        raw_stock: rawStock - count,
        updated_at: new Date().toISOString(),
      })
      .eq('id', product_id);

    if (updateErr) throw updateErr;

    // 아임웹 재고 동기화 (stock_quantity는 변하지 않으므로 동기화 불필요하지만 안전을 위해)
    // stock_quantity = raw_stock + 시리얼 수이므로 총합은 동일

    const startStr = serialNumbers[0];
    const endStr = serialNumbers[serialNumbers.length - 1];

    return NextResponse.json({
      created: count,
      range: `${startStr} ~ ${endStr}`,
      first_serial: startStr,
      raw_stock_remaining: rawStock - count,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
