import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/customers — 고객 신규 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { name, phone, email, address, postcode, addressDetail } = body as {
      name: string;
      phone?: string;
      email?: string;
      address?: string;
      postcode?: string;
      addressDetail?: string;
    };

    if (!name?.trim()) {
      return NextResponse.json({ error: '고객명은 필수입니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: customer, error } = await db
      .from('customers')
      .insert({
        name: name.trim(),
        phone: phone || null,
        email: email || null,
        postcode: postcode || null,
        address_road: address || null,
        address_detail: addressDetail || null,
        source: 'manual',
      })
      .select('id, name, phone, email, address_road, address_detail, postcode, ecount_customer_code, source')
      .single();

    if (error) throw error;

    return NextResponse.json({ customer });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
