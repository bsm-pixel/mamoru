import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/import/customers — 이카운트 고객 CSV 업로드
 *
 * 예상 컬럼: 이름, 영업단가그룹명, 모바일연락처, 매장명, 주소
 * 영업단가그룹명 매핑: 딜러/도매 → dealer, 아카데미/교육 → academy, 그 외 → retail
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const rows: Array<Record<string, string>> = body.rows;

    if (!rows || rows.length === 0) {
      return NextResponse.json({ error: '데이터가 없습니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    let created = 0;
    let skipped = 0;
    const errors: string[] = [];

    for (const row of rows) {
      try {
        const name = (row['이름'] || row['거래처명'] || '').trim();
        if (!name) { skipped++; continue; }

        const phone = (row['모바일연락처'] || row['전화번호'] || row['연락처'] || '').trim().replace(/[^0-9-]/g, '');
        const groupName = (row['영업단가그룹명'] || row['단가그룹'] || '').trim();
        const shopName = (row['매장명'] || row['상호'] || '').trim();
        const address = (row['주소'] || '').trim();

        // B2B 매핑 — 딜러/아카데미 세분화
        function mapCustomerType(g: string): string {
          if (g.includes('딜러') || g.includes('도매')) return 'dealer';
          if (g.includes('아카데미') || g.includes('교육')) return 'academy';
          if (g.startsWith('B2B')) return 'dealer'; // B2B 기본값은 딜러
          return 'retail';
        }
        const customerType = mapCustomerType(groupName);

        // 중복 체크 (이름 + 전화번호)
        let query = db.from('customers').select('id').eq('name', name);
        if (phone) query = query.eq('phone', phone);
        const { data: existing } = await query.limit(1);

        if (existing && existing.length > 0) {
          skipped++;
          continue;
        }

        await db.from('customers').insert({
          name,
          phone: phone || null,
          customer_type: customerType,
          company_name: shopName || null,
          address_road: address || null,
          source: 'manual',
          memo: groupName ? `이카운트: ${groupName}` : null,
          created_at: new Date().toISOString(),
        });
        created++;
      } catch (err) {
        errors.push(`${row['이름'] || '?'}: ${String(err)}`);
      }
    }

    return NextResponse.json({ created, skipped, errors, total: rows.length });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
