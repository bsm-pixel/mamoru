import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { epostSecretsFromEnv, launchEpostBrowser, runEpostPickup, validateRepairForEpost, type EpostRepair } from '@/lib/epost/pickup';

// 우체국 방문접수 자동화 — 서울 리전에서 실행(해외 IP 차단 회피), 크롬 구동 시간 확보
export const runtime = 'nodejs';
export const preferredRegion = 'icn1';
export const maxDuration = 120;

const LOCK_MS = 3 * 60 * 1000; // 진행 중 잠금 3분 — 두 번 클릭·다른 기기 동시 요청 거절
const SELECT = 'id, as_id, name, phone, proceed_type, postcode, address, address_detail, pickup_date, delivery_method, status, epost_reservation_no, epost_requested_at';

type Row = EpostRepair & { id: string; proceed_type: string | null; status: string; epost_reservation_no: string | null; epost_requested_at: string | null };

async function authed() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return { user, db: supabase as any };
}

/**
 * GET /api/repair/epost-pickup?as=AS-... — 시험 전용: '신청' 직전까지(로그인·입력·주소검증·카드검증)만.
 *   주소창 미리 불러오기로 실수 접수되지 않게 GET 은 절대 신청하지 않고, DB 도 바꾸지 않는다.
 */
export async function GET(req: NextRequest) {
  const asId = req.nextUrl.searchParams.get('as') || '';
  if (!/^AS-\d{8}-\d{4}$/.test(asId)) return NextResponse.json({ error: 'as=AS-YYYYMMDD-NNNN 필요' }, { status: 400 });
  const { user, db } = await authed();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  const secrets = epostSecretsFromEnv();
  if (!secrets) return NextResponse.json({ error: '서버에 우체국 로그인·카드 환경변수(EPOST_*)가 없음' }, { status: 500 });
  const { data: rep } = await db.from('repairs').select(SELECT).eq('as_id', asId).maybeSingle();
  if (!rep) return NextResponse.json({ error: `${asId} 없음` }, { status: 404 });
  const bad = validateRepairForEpost(rep);
  if (bad) return NextResponse.json({ error: bad }, { status: 400 });
  const browser = await launchEpostBrowser();
  try {
    const result = await runEpostPickup({ browser, repair: rep as Row, secrets, submit: false });
    return NextResponse.json({ as_id: asId, region: process.env.VERCEL_REGION || 'local', submit: false, ...result }, { status: result.ok ? 200 : 422 });
  } finally {
    await browser.close().catch(() => {});
  }
}

/**
 * POST /api/repair/epost-pickup { repair_id } — 실제 접수 (복원수리 상세 [우체국 수거접수])
 *   1) 방문수거·신규접수·예약번호 없음·진행 중 아님 → 잠금(조건부 업데이트)
 *   2) 서버 크롬으로 우체국 접수 (신청까지)
 *   3) 성공: 예약번호 저장 + 상태 신규접수 → 입고대기 + 이력 / 실패: 잠금 해제(신청 전 실패만)
 *   신청은 눌렀는데 결과를 못 읽으면 '확인필요' 로 막는다 — 다시 누르면 중복 접수될 수 있음
 */
export async function POST(req: NextRequest) {
  const body = await req.json().catch(() => ({}));
  const repairId = String(body.repair_id || '');
  if (!repairId) return NextResponse.json({ error: 'repair_id 필요' }, { status: 400 });
  const { user, db } = await authed();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  const secrets = epostSecretsFromEnv();
  if (!secrets) return NextResponse.json({ error: '서버에 우체국 로그인·카드 환경변수(EPOST_*)가 없음' }, { status: 500 });

  const { data: rep, error: getErr } = await db.from('repairs').select(SELECT).eq('id', repairId).maybeSingle();
  if (getErr) return NextResponse.json({ error: getErr.message }, { status: 500 });
  if (!rep) return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
  const r = rep as Row;
  if (r.status !== 'intake') return NextResponse.json({ error: '신규접수 상태에서만 우체국 접수할 수 있습니다' }, { status: 400 });
  if (r.epost_reservation_no) return NextResponse.json({ error: `이미 우체국 접수된 건입니다 (${r.epost_reservation_no})` }, { status: 409 });
  const bad = validateRepairForEpost(r);
  if (bad) return NextResponse.json({ error: bad }, { status: 400 });

  // 잠금 — 진행 중이 아닐 때만(3분 지난 잠금은 만료로 본다). 조건부 업데이트라 동시에 눌러도 하나만 통과
  const now = new Date();
  const staleBefore = new Date(now.getTime() - LOCK_MS).toISOString();
  const { data: locked, error: lockErr } = await db.from('repairs')
    .update({ epost_requested_at: now.toISOString() })
    .eq('id', repairId).eq('status', 'intake').is('epost_reservation_no', null)
    .or(`epost_requested_at.is.null,epost_requested_at.lt.${staleBefore}`)
    .select('id');
  if (lockErr) return NextResponse.json({ error: lockErr.message }, { status: 500 });
  if (!locked || locked.length === 0) return NextResponse.json({ error: '이미 우체국 접수가 진행 중입니다. 잠시 후 다시 확인해 주세요.' }, { status: 409 });

  const browser = await launchEpostBrowser().catch((e) => e as Error);
  if (browser instanceof Error) {
    await db.from('repairs').update({ epost_requested_at: null }).eq('id', repairId);
    return NextResponse.json({ error: `서버 브라우저 실행 실패: ${browser.message}` }, { status: 500 });
  }
  try {
    const result = await runEpostPickup({ browser, repair: r, secrets, submit: true });
    const booked = result.submitted || !!result.reservationNo;

    if (booked) {
      const no = result.reservationNo || '확인필요';
      await db.from('repairs').update({ epost_reservation_no: no, status: 'pickup_scheduled' }).eq('id', repairId);
      await db.from('repair_history').insert({
        repair_id: repairId, from_status: 'intake', to_status: 'pickup_scheduled', changed_by: user.id,
        note: `우체국 방문접수 자동 예약 ${no} (수거 ${r.pickup_date})`,
      });
    } else if (result.submitAttempted) {
      // 신청은 눌렀으나 예약 여부를 못 읽음 → 재접수 막고 사람 확인
      await db.from('repairs').update({ epost_reservation_no: '확인필요' }).eq('id', repairId);
      await db.from('repair_history').insert({
        repair_id: repairId, from_status: r.status, to_status: r.status, changed_by: user.id,
        note: `우체국 자동접수: 신청 후 결과 확인 실패 — 우체국 예약조회에서 직접 확인 필요 (${result.stoppedReason || ''})`,
      });
    } else {
      // 신청 전에 멈춤 → 아무것도 접수되지 않음. 잠금만 푼다
      await db.from('repairs').update({ epost_requested_at: null }).eq('id', repairId);
    }

    return NextResponse.json({
      ok: booked, reservationNo: booked ? (result.reservationNo || '확인필요') : null,
      needsCheck: !booked && result.submitAttempted,
      stoppedReason: result.stoppedReason, logs: result.logs, ms: result.ms,
      region: process.env.VERCEL_REGION || 'local',
    }, { status: booked ? 200 : 422 });
  } finally {
    await (browser as import('puppeteer-core').Browser).close().catch(() => {});
  }
}
