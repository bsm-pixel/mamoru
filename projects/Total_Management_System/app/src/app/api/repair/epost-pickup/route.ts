import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { epostSecretsFromEnv, launchEpostBrowser, runEpostPickup, validateRepairForEpost, type EpostRepair } from '@/lib/epost/pickup';

// 우체국 방문접수 자동화 — 서울 리전에서 실행(해외 IP 차단 회피), 크롬 구동 시간 확보
export const runtime = 'nodejs';
export const preferredRegion = 'icn1';
export const maxDuration = 120;

async function load(asId: string) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return { err: NextResponse.json({ error: 'Unauthorized' }, { status: 401 }) };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = supabase as any;
  const { data: rep, error } = await db
    .from('repairs')
    .select('id, as_id, name, phone, proceed_type, postcode, address, address_detail, pickup_date, delivery_method, status')
    .eq('as_id', asId)
    .maybeSingle();
  if (error) return { err: NextResponse.json({ error: error.message }, { status: 500 }) };
  if (!rep) return { err: NextResponse.json({ error: `${asId} 없음` }, { status: 404 }) };
  const bad = validateRepairForEpost(rep);
  if (bad) return { err: NextResponse.json({ error: bad }, { status: 400 }) };
  return { rep: rep as EpostRepair & { id: string } };
}

async function run(asId: string, submit: boolean) {
  const secrets = epostSecretsFromEnv();
  if (!secrets) return NextResponse.json({ error: '서버에 우체국 로그인·카드 환경변수(EPOST_*)가 없음' }, { status: 500 });
  const { rep, err } = await load(asId);
  if (err) return err;
  const browser = await launchEpostBrowser();
  try {
    const result = await runEpostPickup({ browser, repair: rep!, secrets, submit });
    return NextResponse.json({ as_id: asId, region: process.env.VERCEL_REGION || 'local', submit, ...result }, { status: result.ok ? 200 : 422 });
  } finally {
    await browser.close().catch(() => {});
  }
}

/**
 * GET /api/repair/epost-pickup?as=AS-... — 시험 전용: '신청' 직전까지(로그인·입력·주소검증·카드검증)만.
 *   주소창 미리 불러오기로 실수 접수되지 않게 GET 은 절대 신청하지 않는다.
 */
export async function GET(req: NextRequest) {
  const asId = req.nextUrl.searchParams.get('as') || '';
  if (!/^AS-\d{8}-\d{4}$/.test(asId)) return NextResponse.json({ error: 'as=AS-YYYYMMDD-NNNN 필요' }, { status: 400 });
  return run(asId, false);
}

/** POST /api/repair/epost-pickup { as_id, submit } — 실제 신청은 여기서만 (3단계 버튼) */
export async function POST(req: NextRequest) {
  const body = await req.json().catch(() => ({}));
  const asId = String(body.as_id || '');
  if (!/^AS-\d{8}-\d{4}$/.test(asId)) return NextResponse.json({ error: 'as_id 필요' }, { status: 400 });
  return run(asId, body.submit === true);
}
