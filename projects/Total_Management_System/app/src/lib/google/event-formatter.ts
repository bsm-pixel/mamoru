/**
 * Consultation → Google Calendar Event 변환기
 * 이벤트 제목/설명/색상/확장프로퍼티 포맷 표준화
 */

import type { calendar_v3 } from 'googleapis';
import { SCHEDULE_COLORS } from '@/lib/schedule/colors';

export interface ConsultationForCalendar {
  id: string;
  name: string | null;
  phone: string | null;
  consultation_type: string;
  visit_date: string | null;
  visit_time: string | null;
  status: string;
  address_road?: string | null;
  address_detail?: string | null;
  address_sigungu?: string | null;
  memo?: string | null;
  adminNote?: string | null;   // 108: 상담자(관리자) 전용 메모 — 캘린더 설명란 반영
  unique_id?: string | null;
  created_at?: string | null;
  completed_at?: string | null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  gas_raw?: any;
}

export interface EventFormatSettings {
  store_name?: string;
  store_address?: string;
  duration_min?: number;
  field_buffer_before?: number;
  field_buffer_after?: number;
  duration_store_visit?: number;
  duration_field_request?: number;
}


function getColorId(type: string, status: string): string {
  // 색상 SSOT(lib/schedule/colors.ts) 참조 — 인앱 일정 달력과 동일 색
  //   매장=초록(emerald→Sage) / 출장=보라(violet→Grape) / 수리=주황(amber→Tangerine)
  if (status === 'reschedule_requested' || status === 'change_requested') return '5'; // Banana (노랑) — 고객 변경요청
  if (status === 'completed') return '8'; // Graphite (회색)
  if (type === 'field_request') return SCHEDULE_COLORS.field.googleColorId ?? '3';
  return SCHEDULE_COLORS.store.googleColorId ?? '2'; // store_visit 기본
}

function toKSTIso(date: string, time: string): string {
  // "2026-04-26" + "14:00" → "2026-04-26T14:00:00+09:00"
  const t = time.length === 5 ? `${time}:00` : time;
  return `${date}T${t}+09:00`;
}

function addMinutes(time: string, minutes: number): string {
  const [h, m] = time.split(':').map(Number);
  const total = h * 60 + m + minutes;
  const nh = Math.floor(total / 60) % 24;
  const nm = total % 60;
  return `${String(nh).padStart(2, '0')}:${String(nm).padStart(2, '0')}`;
}

function formatDateKR(iso?: string | null): string {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return iso;
    // KST 변환
    const kst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
    const y = kst.getUTCFullYear();
    const mo = String(kst.getUTCMonth() + 1).padStart(2, '0');
    const da = String(kst.getUTCDate()).padStart(2, '0');
    const hh = String(kst.getUTCHours()).padStart(2, '0');
    const mm = String(kst.getUTCMinutes()).padStart(2, '0');
    return `${y}-${mo}-${da} ${hh}:${mm}`;
  } catch {
    return iso;
  }
}

/**
 * 제목에 쓸 지역 한 토막 — '서초구' · '일산동구'
 * address_sigungu 는 실측 전부 NULL(2026-09-24) 이라 도로명 주소에서 뽑는다.
 * '경기 고양시 일산동구 …' 처럼 시·구가 겹치면 더 좁은 쪽(구)을 쓴다.
 */
function shortRegion(c: ConsultationForCalendar): string {
  if (c.address_sigungu) return c.address_sigungu;
  const tokens = (c.address_road || '').trim().split(/\s+/).slice(0, 3);
  const hits = tokens.filter((t) => /(시|군|구)$/.test(t));
  return hits.length ? hits[hits.length - 1] : '';
}

/** 사람이 못 읽는 UUID 형태면 본문에 넣지 않는다 (상담번호가 UUID 인 건이 있음) */
function readableId(id?: string | null): string | null {
  if (!id) return null;
  return /^[0-9a-f]{8}-[0-9a-f]{4}-/i.test(id) ? null : id;
}

function buildFullAddress(c: ConsultationForCalendar): string {
  const road = (c.address_road || '').trim();
  const detail = (c.address_detail || '').trim();
  if (road && detail) return `${road} ${detail}`;
  return road || detail || '';
}

/** 출장 이벤트 기본 소요 시간 (분) */
function getDurationMin(type: string, settings: EventFormatSettings): number {
  if (type === 'field_request') return settings.duration_field_request ?? settings.duration_min ?? 60;
  return settings.duration_store_visit ?? settings.duration_min ?? 60;
}

/** 짧은 날짜 — '9/22 14:03' (KST) */
function shortDateTimeKR(iso?: string | null): string {
  const full = formatDateKR(iso);            // 'YYYY-MM-DD HH:MM'
  if (!full || full.length < 16) return full;
  return `${Number(full.slice(5, 7))}/${Number(full.slice(8, 10))} ${full.slice(11, 16)}`;
}

/**
 * 설명 블록 표준 (2026-09-24 정리)
 *   - 구분선(━)·항목마다 붙던 이모지·3줄 경고문 전부 제거. 캘린더에서 **읽을 게 아니라 확인할 것**만 남긴다.
 *   - 순서 = 위계: ① 연락 ② 이 건의 내용 ③ 식별·출처 ④ TMS 링크
 *   - 주소는 location 필드에 들어가므로 본문에서 반복하지 않는다(구글이 지도·길찾기로 띄워줌).
 */
function buildDescription(blocks: Array<string | null | undefined>, baseUrl: string, path: string): string {
  const body = blocks.map((b) => (b || '').trim()).filter(Boolean);
  body.push(`${baseUrl}${path}`);
  return body.join('\n');
}

/**
 * Consultation → Calendar Event 변환
 *
 * 제목 규칙: `종류 · 이름 · 지역` — 한눈에 확인할 것만. 전화번호는 제목에서 뺀다(잘림 유발, 본문에 있음)
 *            확인이 필요한 상태만 앞에 [변경요청] 을 붙인다.
 */
export function formatConsultationToEvent(
  c: ConsultationForCalendar,
  settings: EventFormatSettings,
  baseUrl: string,
): calendar_v3.Schema$Event {
  const isField = c.consultation_type === 'field_request';
  const typeLabel = isField ? '출장' : '매장';
  const region = shortRegion(c);
  const name = c.name || '고객';
  const phone = c.phone || '';
  const durMin = getDurationMin(c.consultation_type, settings);
  const needsAttention = c.status === 'reschedule_requested' || c.status === 'change_requested';

  const summary = [
    needsAttention ? '[변경요청]' : '',
    [typeLabel, name, isField && region ? region : ''].filter(Boolean).join(' · '),
  ].filter(Boolean).join(' ');

  const fullAddress = buildFullAddress(c);
  const location = isField ? fullAddress : settings.store_address || '';

  const reschedReason =
    c.gas_raw?.reschedule_reason ||
    c.gas_raw?.change_reason ||
    c.gas_raw?.rescheduleReason;

  const description = buildDescription([
    phone,
    needsAttention && reschedReason ? `변경 요청: ${reschedReason}` : null,
    c.memo ? `고객 메모: ${c.memo}` : null,
    c.adminNote ? `내 메모: ${c.adminNote}` : null,
    [readableId(c.unique_id), c.created_at ? `접수 ${shortDateTimeKR(c.created_at)}` : null]
      .filter(Boolean).join(' · ') || null,
  ], baseUrl, `/consultations/${c.id}`);

  const startTime = c.visit_time && c.visit_time.match(/^\d{1,2}:\d{2}/) ? c.visit_time.slice(0, 5) : '10:00';
  const endTime = addMinutes(startTime, durMin);
  const visitDate = c.visit_date || '';

  const event: calendar_v3.Schema$Event = {
    summary,
    description,
    location: location || undefined,
    start: visitDate
      ? { dateTime: toKSTIso(visitDate, startTime), timeZone: 'Asia/Seoul' }
      : undefined,
    end: visitDate
      ? { dateTime: toKSTIso(visitDate, endTime), timeZone: 'Asia/Seoul' }
      : undefined,
    colorId: getColorId(c.consultation_type, c.status),
    // 기본 리마인더 OFF — 알림톡·푸시 중복 방지 (설정 UI에서 ON 가능 — 추후)
    reminders: { useDefault: false, overrides: [] },
    // 숨김 메타데이터 — 역동기화/분쟁 해결용
    extendedProperties: {
      private: {
        mamoru_consultation_id: c.id,
        mamoru_consultation_type: c.consultation_type,
        mamoru_status: c.status,
        mamoru_version: '2.0',
      },
    },
  };

  return event;
}

// ═══════════════════════════════════════════════════════════════════
// 2026-05-25 Phase 3-B: 복원수리 직접방문(당일수리) → Google Calendar
// 컨설팅 패턴 동일 (재사용) — repairs.proceed_type='직접방문' 전용
// ═══════════════════════════════════════════════════════════════════

export interface RepairForCalendar {
  id: string;
  as_id: string;
  name: string | null;
  phone: string | null;
  visit_date: string | null;
  visit_time: string | null;
  visit_duration_min: number | null;
  status: string;
  qty_mamoru: number | null;
  qty_other: number | null;
  memo?: string | null;
  service_cost?: number | null;
  total_amount?: number | null;
  created_at?: string | null;
}

function getRepairColorId(status: string): string {
  // 복원수리 직접방문 색상 — 색상 SSOT 참조 (인앱 수리=amber→Tangerine)
  if (status === 'cancelled') return '8';      // Graphite (회색)
  if (status === 'completed') return '8';      // Graphite (회색) — 완료
  return SCHEDULE_COLORS.repair_visit.googleColorId ?? '6'; // Tangerine (주황)
}

/**
 * 복원수리 직접방문 → Calendar Event 변환
 * 컨설팅 패턴 동일 (재사용)
 */
export function formatRepairToEvent(
  r: RepairForCalendar,
  settings: EventFormatSettings,
  baseUrl: string,
): calendar_v3.Schema$Event {
  const name = r.name || '고객';
  const phone = r.phone || '';
  const qtyM = r.qty_mamoru || 0;
  const qtyO = r.qty_other || 0;
  const qty = qtyM + qtyO;
  const durMin = r.visit_duration_min || (qty >= 6 ? 60 : 30);

  // 제목: `수리 · 고객명 · N자루`
  const summary = ['수리', name, qty > 0 ? `${qty}자루` : ''].filter(Boolean).join(' · ');

  const description = buildDescription([
    phone,
    qty > 0 ? `마모루 ${qtyM}자루 · 타사 ${qtyO}자루 · 예상 ${durMin}분` : `예상 ${durMin}분`,
    r.memo ? `고객 메모: ${r.memo}` : null,
    [r.as_id || null, r.created_at ? `접수 ${shortDateTimeKR(r.created_at)}` : null]
      .filter(Boolean).join(' · ') || null,
  ], baseUrl, `/repairs/${r.id}`);

  const startTime = r.visit_time && r.visit_time.match(/^\d{1,2}:\d{2}/) ? r.visit_time.slice(0, 5) : '10:00';
  const endTime = addMinutes(startTime, durMin);
  const visitDate = r.visit_date || '';

  const event: calendar_v3.Schema$Event = {
    summary,
    description,
    location: settings.store_address || undefined,   // 직접방문 = 매장 워크인
    start: visitDate
      ? { dateTime: toKSTIso(visitDate, startTime), timeZone: 'Asia/Seoul' }
      : undefined,
    end: visitDate
      ? { dateTime: toKSTIso(visitDate, endTime), timeZone: 'Asia/Seoul' }
      : undefined,
    colorId: getRepairColorId(r.status),
    reminders: { useDefault: false, overrides: [] },
    extendedProperties: {
      private: {
        mamoru_repair_id: r.id,
        mamoru_repair_as_id: r.as_id,
        mamoru_repair_status: r.status,
        mamoru_source: 'repair_direct_visit',
        mamoru_version: '2.0',
      },
    },
  };

  return event;
}

/**
 * 송장 미생성 '할 일' → 종일 이벤트 (2026-09-24)
 * 같은 제목·본문 규칙을 따른다: `송장 생성 · 이름` / 본문 3줄 이내.
 */
export function formatShippingTodoToEvent(t: {
  kind: 'sale' | 'delivery';
  who: string;
  docNo: string | null;
  amount: number | null;
  createdAt: string | null;
  date: string;        // 표시할 날짜(KST, 종일)
  nextDate: string;    // 종료일(배타적)
}, baseUrl: string): calendar_v3.Schema$Event {
  const isSale = t.kind === 'sale';
  const summary = `${isSale ? '송장 생성' : '납품 출고'} · ${t.who}`;
  const description = buildDescription([
    [t.docNo || null, t.amount ? `${Number(t.amount).toLocaleString('ko-KR')}원` : null]
      .filter(Boolean).join(' · ') || null,
    t.createdAt ? `등록 ${shortDateTimeKR(t.createdAt)} · 송장 없음` : '송장 없음',
  ], baseUrl, isSale ? '/sales' : '/deliveries');

  return {
    summary,
    description,
    start: { date: t.date },
    end: { date: t.nextDate },          // 구글 종일 일정의 end.date 는 배타적
    colorId: '8',                        // Graphite — 예약(초록·보라·주황)과 구분되는 '업무'
    transparency: 'transparent',         // 한가함 — 예약 슬롯을 막지 않는다
    reminders: { useDefault: false, overrides: [] },
    extendedProperties: { private: { mamoru_type: 'shipping_todo', mamoru_ref: `${t.kind}:${t.docNo || ''}` } },
  };
}
