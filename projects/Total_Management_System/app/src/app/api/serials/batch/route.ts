import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/serials/batch — 시리얼 일괄 생성 */
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

    // 기존 시리얼 수 조회 (넘버링 이어서)
    const { count: existing } = await db
      .from('product_serials')
      .select('*', { count: 'exact', head: true })
      .like('serial_number', `${basePrefix}%`);

    const startNum = (existing || 0) + 1;

    const serials = Array.from({ length: count }, (_, i) => ({
      product_id,
      serial_number: `${basePrefix}-${String(startNum + i).padStart(3, '0')}`,
      lot_number: lot_number || null,
      created_by: user.id,
    }));

    const { error } = await db
      .from('product_serials')
      .insert(serials);

    if (error) throw error;

    return NextResponse.json({ created: count, prefix: basePrefix });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
