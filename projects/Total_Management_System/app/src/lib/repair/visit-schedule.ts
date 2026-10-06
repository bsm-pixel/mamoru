/**
 * 직접방문(매장 당일수리) 시간표 — 공용 (2026-10-06)
 *
 * 고객 접수 페이지 슬롯(api/repair/public/slots)과 TMS 「방문 예약 등록」 충돌 경고가 같은 계산을 쓴다.
 * (원래 public/slots 안에 있던 로직을 옮김 — 차단 범위 규칙은 그대로. 각 차단에 사람이 읽는 라벨만 붙였다)
 *
 * 2026-10-06 같이 고친 것: "오늘 지난 시간" 판정이 서버(UTC) 시각 기준이라 KST 오전엔 이미 지난 시간이
 * 예약 가능으로 보였다 → KST 기준으로 계산.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface BusyRange { start: number; end: number; label: string }

export function toMinutes(hhmm: string): number {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + (m || 0);
}
export function fromMinutes(min: number): string {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}
export function overlaps(a0: number, a1: number, b0: number, b1: number): boolean {
  return a0 < b1 && b0 < a1;
}
const hhmm = (t: string) => String(t || '').slice(0, 5);

/** 그 날짜의 영업 설정 + 휴무 여부 + 점유(차단) 구간 */
export async function getVisitDay(db: Db, date: string) {
  const { data: settings } = await db.from('consultation_settings').select('*').eq('id', 'default').single();

  const cfg = {
    startHour: (settings?.start_hour ?? 10) as number,
    endHour: (settings?.end_hour ?? 20) as number,
    fieldBufferBefore: (settings?.field_buffer_before ?? 60) as number,
    fieldBufferAfter: (settings?.field_buffer_after ?? 60) as number,
    consultDurMin: (settings?.duration_min ?? 60) as number,       // 매장방문/출장 차단 길이
    disabledWeekdays: (settings?.disabled_weekdays ?? [0]) as number[],
    slotStep: (settings?.repair_slot_step_min ?? 30) as number,     // 092 슬롯 간격
  };

  // 휴무 (요일 / 지정 휴무일) — 날짜 문자열만 쓰므로 서버 타임존 무관하게 UTC 요일로 계산
  const [y, mo, d] = date.split('-').map(Number);
  const dayOfWeek = new Date(Date.UTC(y, mo - 1, d)).getUTCDay();
  let closedReason: 'disabled_weekday' | 'closed_date' | null = cfg.disabledWeekdays.includes(dayOfWeek) ? 'disabled_weekday' : null;
  if (!closedReason) {
    const { data: closedDate } = await db.from('closed_dates').select('date').eq('date', date).maybeSingle();
    if (closedDate) closedReason = 'closed_date';
  }

  // 같은 날 충돌 데이터 + 시간대 차단(096 날짜 / 118 매주반복) 병렬 조회
  const [consultsRes, suggestedRes, repairsRes, blockedRes, weeklyBlockedRes] = await Promise.all([
    // 컨설팅 매장방문 + 출장 — 확정/배정/접수대기까지 점유
    db.from('consultations')
      .select('consultation_type, visit_time, status, name')
      .eq('visit_date', date)
      .in('status', ['confirmed', 'assigned', 'pending_admin']),
    // 제안(suggested) 건: suggestions JSONB 안의 날짜로 매칭
    db.from('consultations')
      .select('consultation_type, suggestions, status, name')
      .eq('status', 'suggested')
      .not('suggestions', 'is', null),
    // 복원수리 직접방문 (취소 제외)
    db.from('repairs')
      .select('visit_time, visit_duration_min, status, name, as_id')
      .eq('visit_date', date)
      .eq('proceed_type', '직접방문')
      .neq('status', 'cancelled'),
    db.from('blocked_time_slots').select('start_time, end_time').eq('date', date),
    db.from('blocked_time_slots').select('weekday, start_time, end_time').eq('weekday', dayOfWeek),
  ]);

  const busy: BusyRange[] = [];
  for (const bs of [...(blockedRes.data || []), ...(weeklyBlockedRes.data || [])]) {
    if (!bs.start_time || !bs.end_time) continue;
    busy.push({ start: toMinutes(bs.start_time), end: toMinutes(bs.end_time), label: `일정 차단 ${hhmm(bs.start_time)}~${hhmm(bs.end_time)}` });
  }
  for (const c of (consultsRes.data || [])) {
    if (!c.visit_time) continue;
    const t = toMinutes(c.visit_time);
    if (c.consultation_type === 'field_request') {
      busy.push({ start: t - cfg.fieldBufferBefore, end: t + cfg.consultDurMin + cfg.fieldBufferAfter, label: `출장 상담 ${hhmm(c.visit_time)} · ${c.name || ''}`.trim() });
    } else {
      busy.push({ start: t, end: t + cfg.consultDurMin, label: `매장 상담 ${hhmm(c.visit_time)} · ${c.name || ''}`.trim() });
    }
  }
  for (const c of (suggestedRes.data || [])) {
    const raw = c.suggestions as { dates?: Array<{ date?: string; time?: string }> } | Array<{ date?: string; time?: string }> | null;
    if (!raw) continue;
    const sug = Array.isArray(raw) ? raw : (raw.dates || []);
    for (const s of sug) {
      if (s.date !== date || !s.time) continue;
      const t = toMinutes(s.time);
      const label = `상담 제안 시간 ${hhmm(s.time)} · ${c.name || ''}`.trim();
      if (c.consultation_type === 'field_request') busy.push({ start: t - cfg.fieldBufferBefore, end: t + cfg.consultDurMin + cfg.fieldBufferAfter, label });
      else busy.push({ start: t, end: t + cfg.consultDurMin, label });
    }
  }
  for (const r of (repairsRes.data || [])) {
    if (!r.visit_time) continue;
    const t = toMinutes(r.visit_time);
    busy.push({ start: t, end: t + ((r.visit_duration_min as number) || 30), label: `복원수리 방문 ${hhmm(r.visit_time)} · ${r.name || ''}`.trim() });
  }

  return { cfg, closedReason, busy };
}

/** 지금 KST 날짜(YYYY-MM-DD)와 분 */
export function kstNow(): { date: string; minutes: number } {
  const k = new Date(Date.now() + 9 * 3600 * 1000);
  return { date: k.toISOString().slice(0, 10), minutes: k.getUTCHours() * 60 + k.getUTCMinutes() };
}

/** 영업시간 안 슬롯 목록 (고객 페이지와 같은 규칙) */
export function buildSlots(day: Awaited<ReturnType<typeof getVisitDay>>, date: string, blockMin: number) {
  const { cfg, busy } = day;
  const startMin = cfg.startHour * 60;
  const endMin = cfg.endHour * 60;
  const now = kstNow();
  const isToday = date === now.date;
  const slots: Array<{ time: string; available: boolean; past?: boolean; conflicts?: string[] }> = [];
  for (let m = startMin; m + blockMin <= endMin; m += cfg.slotStep) {
    if (isToday && m <= now.minutes) { slots.push({ time: fromMinutes(m), available: false, past: true }); continue; }
    const end = m + blockMin;
    const hits = busy.filter((b) => overlaps(m, end, b.start, b.end)).map((b) => b.label);
    slots.push({ time: fromMinutes(m), available: hits.length === 0, ...(hits.length ? { conflicts: hits } : {}) });
  }
  return slots;
}

/** 특정 시각에 겹치는 일정 라벨 (영업시간 밖이어도 겹침만 본다) */
export function conflictsAt(day: Awaited<ReturnType<typeof getVisitDay>>, time: string, blockMin: number): string[] {
  const s = toMinutes(time), e = s + blockMin;
  return day.busy.filter((b) => overlaps(s, e, b.start, b.end)).map((b) => b.label);
}
