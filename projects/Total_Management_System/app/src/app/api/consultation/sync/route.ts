import { NextResponse } from 'next/server';
import { syncConsultations } from '@/lib/consultation/sync';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/consultation/sync — GAS 상담 동기화 */
export async function POST() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const result = await syncConsultations();
    return NextResponse.json(result);
  } catch (err) {
    console.error('[consultation-sync] 동기화 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500 }
    );
  }
}
