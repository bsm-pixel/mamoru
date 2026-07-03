import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getOpenApiConnectionStatus } from '@/lib/imweb/client';

/** GET /api/imweb/oauth/status — 아임웹 OpenAPI 연결 상태 진단 (재고 push 가능 여부) */
export async function GET() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const status = await getOpenApiConnectionStatus();
    return NextResponse.json(status);
  } catch (err) {
    return NextResponse.json({ connected: false, error: String(err) }, { status: 200 });
  }
}
