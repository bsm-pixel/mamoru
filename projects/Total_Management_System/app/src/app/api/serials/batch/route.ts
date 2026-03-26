import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** 혼동 문자 제외 랜덤 4자리 생성 (0/O, 1/I/L 제외) */
function generateRandomSuffix(): string {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  return Array.from({ length: 4 }, () => chars[Math.floor(Math.random() * chars.length)]).join('');
}

/** POST /api/serials/batch — 시리얼 일괄 생성 (랜덤 4자리) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { product_id, count, prefix, lot_number } = await req.json() as {
      product_id: string;
      count: number;
      prefix?: string;
      lot_number?: string;
    };

    if (!product_id || !count || count < 1 || count > 100) {
      return NextResponse.json({ error: '1~100개 범위로 입력해주세요' }, { status: 400 });
    }

    // 제품 SKU 조회
    const { data: product } = await db
      .from('products')
      .select('sku')
      .eq('id', product_id)
      .single();

    const sku = product?.sku || 'XX';
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const basePrefix = prefix || `MM-${sku}-${dateStr}`;

    // 랜덤 4자리 시리얼 생성 (중복 방지)
    const usedSuffixes = new Set<string>();
    const serials = [];

    for (let i = 0; i < count; i++) {
      let suffix: string;
      let serialNumber: string;
      let attempts = 0;
      do {
        suffix = generateRandomSuffix();
        serialNumber = `${basePrefix}-${suffix}`;
        attempts++;
        if (attempts > 100) throw new Error('시리얼 생성 실패: 중복 초과');
      } while (usedSuffixes.has(suffix));
      usedSuffixes.add(suffix);

      serials.push({
        product_id,
        serial_number: serialNumber,
        barcode: serialNumber, // QR 출력용 barcode 자동 채움
        lot_number: lot_number || null,
        created_by: user.id,
      });
    }

    // DB UNIQUE 제약이 최종 안전장치
    const { error } = await db
      .from('product_serials')
      .insert(serials);

    if (error) throw error;

    return NextResponse.json({ created: count, prefix: basePrefix });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
