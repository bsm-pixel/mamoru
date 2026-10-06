import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { createRepairIntake, visitDurationMin } from '@/lib/repair/intake';
import { getVisitDay, buildSlots, conflictsAt, kstNow, toMinutes } from '@/lib/repair/visit-schedule';

/**
 * TMS 「방문 예약 등록」 — 전화로 "몇 시쯤 갈게요" 한 고객을 사장님이 직접 등록 (2026-10-06)
 *
 *   GET  ?date=YYYY-MM-DD&qty=N  → 그날 시간표 (비어있는 시간 + 겹치는 일정 이름까지 — 관리자 전용)
 *   POST { name, phone, visit_date, visit_time, qty_mamoru, qty_other, memo, force? }
 *        → 고객 접수와 같은 공용 함수(createRepairIntake, actor='admin')로 생성
 *          = 예약번호 · 고객 연결 · 「매장방문 접수」 알림톡(일정 변경 링크) · 구글 캘린더 · 리마인드 대상
 *
 * 충돌 처리 = **경고 후 허용** (사장님 결정): 다른 일정과 겹치거나 휴무일·영업시간 밖·지난 시각이면
 *   force 없이는 409 { needConfirm, warnings } → 화면이 확인창 → force:true 로 다시 보내면 등록.
 *   (고객 페이지는 그대로 겹치는 시간을 아예 못 고른다)
 */

async function authed() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  return user;
}

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const TIME_RE = /^([01]\d|2[0-3]):[0-5]\d$/;

export async function GET(req: NextRequest) {
  try {
    if (!(await authed())) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const date = req.nextUrl.searchParams.get('date') || '';
    const qty = Math.max(1, parseInt(req.nextUrl.searchParams.get('qty') || '1', 10) || 1);
    if (!DATE_RE.test(date)) return NextResponse.json({ error: 'date(YYYY-MM-DD) required' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const blockMin = visitDurationMin(qty);
    const day = await getVisitDay(db, date);
    return NextResponse.json({
      date, qty, blockMin,
      closedReason: day.closedReason,
      businessHours: { start: day.cfg.startHour, end: day.cfg.endHour },
      slots: buildSlots(day, date, blockMin),
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

export async function POST(req: NextRequest) {
  try {
    if (!(await authed())) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const b = await req.json().catch(() => ({}));

    const name = String(b.name || '').trim();
    const phone = String(b.phone || '').trim();
    const visitDate = String(b.visit_date || '');
    const visitTime = String(b.visit_time || '').slice(0, 5);
    const qtyM = Math.max(0, parseInt(b.qty_mamoru, 10) || 0);
    const qtyO = Math.max(0, parseInt(b.qty_other, 10) || 0);

    if (!name) return NextResponse.json({ error: '성함을 입력해주세요' }, { status: 400 });
    if (phone.replace(/\D/g, '').length < 10) return NextResponse.json({ error: '연락처를 확인해주세요' }, { status: 400 });
    if (!DATE_RE.test(visitDate)) return NextResponse.json({ error: '방문 날짜를 선택해주세요' }, { status: 400 });
    if (!TIME_RE.test(visitTime)) return NextResponse.json({ error: '방문 시간을 선택해주세요' }, { status: 400 });
    if (qtyM + qtyO < 1) return NextResponse.json({ error: '가위 수를 1자루 이상 입력해주세요' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;

    // ── 경고 모으기 (막지 않고, 확인을 받는다) ──
    const blockMin = visitDurationMin(qtyM + qtyO);
    const day = await getVisitDay(db, visitDate);
    const warnings: string[] = [];
    if (day.closedReason === 'disabled_weekday') warnings.push('정기 휴무 요일입니다');
    if (day.closedReason === 'closed_date') warnings.push('휴무일로 지정된 날입니다');
    const t = toMinutes(visitTime);
    if (t < day.cfg.startHour * 60 || t + blockMin > day.cfg.endHour * 60) {
      warnings.push(`영업시간(${day.cfg.startHour}시~${day.cfg.endHour}시) 밖입니다`);
    }
    const now = kstNow();
    if (visitDate < now.date || (visitDate === now.date && t <= now.minutes)) warnings.push('이미 지난 시간입니다');
    for (const c of conflictsAt(day, visitTime, blockMin)) warnings.push(`겹치는 일정: ${c}`);

    // 같은 연락처로 같은 날 이미 직접방문 예약이 있으면 중복 등록일 가능성
    const { data: dup } = await db.from('repairs')
      .select('as_id, visit_time')
      .eq('phone_normalized', phone.replace(/\D/g, ''))
      .eq('visit_date', visitDate).eq('proceed_type', '직접방문').neq('status', 'cancelled');
    for (const d of dup || []) warnings.push(`같은 고객의 같은 날 예약이 이미 있습니다 (${d.as_id} · ${String(d.visit_time || '').slice(0, 5)})`);

    if (warnings.length && b.force !== true) {
      return NextResponse.json({ needConfirm: true, warnings }, { status: 409 });
    }

    const memoBase = String(b.memo || '').trim();
    const result = await createRepairIntake(db, {
      name, phone,
      proceed_type: '직접방문',
      visit_date: visitDate,
      visit_time: visitTime,
      qty_mamoru: qtyM,
      qty_other: qtyO,
      memo: [memoBase, warnings.length ? `[확인 후 등록] ${warnings.join(' / ')}` : ''].filter(Boolean).join('\n') || null,
    }, 'admin');

    return NextResponse.json({
      ok: true,
      id: result.repair.id,
      as_id: result.asId,
      notified: result.notified,
      warnings,
    }, { status: 201 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
