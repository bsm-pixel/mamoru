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

const STATUS_KR: Record<string, string> = {
  confirmed: '확정',
  reschedule_requested: '고객 변경 요청 중',
  change_requested: '고객 변경 요청 중',
  completed: '완료',
};

function getStatusPrefix(status: string): string {
  if (status === 'reschedule_requested' || status === 'change_requested') return '⏳ ';
  if (status === 'completed') return '✅ ';
  return '';
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

/**
 * Consultation → Calendar Event 변환
 *
 * @param c 상담 레코드
 * @param settings 매장/기본값 설정
 * @param baseUrl TMS 앱 베이스 URL (이벤트 설명의 링크용)
 */
export function formatConsultationToEvent(
  c: ConsultationForCalendar,
  settings: EventFormatSettings,
  baseUrl: string,
): calendar_v3.Schema$Event {
  const isField = c.consultation_type === 'field_request';
  const typeLabel = isField ? '[출장]' : '[매장]';
  const statusPrefix = getStatusPrefix(c.status);
  const region = c.address_sigungu || '';
  const regionLabel = isField && region ? ` · ${region}` : '';

  const durMin = getDurationMin(c.consultation_type, settings);
  const name = c.name || '고객';
  const phone = c.phone || '';

  // 제목: [매장/출장] ⏳✅ 이름 · 지역 · 010-xxxx
  const summary = `${typeLabel} ${statusPrefix}${name}${regionLabel}${phone ? ' · ' + phone : ''}`.trim();

  // 위치
  const fullAddress = buildFullAddress(c);
  const location = isField ? fullAddress : settings.store_address || '';

  // 설명 (Description) — 풍부한 정보 블록
  const lines: string[] = [];
  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push(`🏷 상담 종류: ${isField ? '출장 요청' : '매장 방문'}`);
  lines.push(`👤 고객명: ${name}`);
  if (phone) lines.push(`📱 연락처: ${phone}`);
  if (isField && fullAddress) lines.push(`📍 방문 주소: ${fullAddress}`);
  if (!isField && settings.store_name) lines.push(`🏪 방문지: ${settings.store_name}`);
  if (c.memo) lines.push(`💬 고객 메모: ${c.memo}`);
  if (c.adminNote) lines.push(`📝 상담자 메모: ${c.adminNote}`);
  if (isField) {
    const before = settings.field_buffer_before ?? 90;
    const after = settings.field_buffer_after ?? 90;
    lines.push(`🚗 이동 버퍼: 앞 ${before}분 / 뒤 ${after}분 (매장 예약 자동 차단)`);
  }
  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push(`📋 상태: ${STATUS_KR[c.status] || c.status}`);
  if (c.unique_id) lines.push(`🆔 상담번호: ${c.unique_id}`);
  if (c.created_at) lines.push(`📅 접수일시: ${formatDateKR(c.created_at)}`);

  // 재요청 사유 (gas_raw에 저장됨)
  const reschedReason =
    c.gas_raw?.reschedule_reason ||
    c.gas_raw?.change_reason ||
    c.gas_raw?.rescheduleReason;
  if ((c.status === 'reschedule_requested' || c.status === 'change_requested') && reschedReason) {
    lines.push(`⚠️ 고객 변경 사유: ${reschedReason}`);
  }

  if (c.status === 'completed' && c.completed_at) {
    lines.push(`✅ 완료 시각: ${formatDateKR(c.completed_at)}`);
  }

  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('');
  lines.push('🔗 TMS 상세보기:');
  lines.push(`${baseUrl}/consultations/${c.id}`);

  if (phone) {
    const phoneDigits = phone.replace(/\D/g, '');
    lines.push('');
    lines.push(`📞 바로 전화: tel:${phoneDigits}`);
  }

  lines.push('');
  lines.push('⚠ 이 일정은 MAMORU TMS에서 자동 생성됩니다.');
  lines.push('  캘린더에서 직접 수정하시면 TMS와 불일치가 발생합니다.');
  lines.push('  변경·취소는 반드시 TMS에서 진행해 주세요.');

  const description = lines.join('\n');

  // 시작/종료 시간
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
        mamoru_version: '1.0',
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
  const qty = (r.qty_mamoru || 0) + (r.qty_other || 0);
  const durMin = r.visit_duration_min || (qty >= 6 ? 60 : 30);

  // 제목: [복원수리 직접방문] 고객명 · N자루 · 010-xxxx
  const qtyLabel = qty > 0 ? ` · ${qty}자루` : '';
  const summary = `[복원수리 직접방문] ${name}${qtyLabel}${phone ? ' · ' + phone : ''}`.trim();

  // 위치: 매장 (직접방문은 매장 워크인)
  const location = settings.store_address || '';

  // Description
  const lines: string[] = [];
  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('🔧 복원수리 (당일수리)');
  lines.push(`👤 고객명: ${name}`);
  if (phone) lines.push(`📱 연락처: ${phone}`);
  if (settings.store_name) lines.push(`🏪 방문지: ${settings.store_name}`);
  if (qty > 0) {
    lines.push(`✂️ 가위 수량: 마모루 ${r.qty_mamoru || 0}자루 / 타사 ${r.qty_other || 0}자루 (총 ${qty}자루)`);
  }
  lines.push(`⏱ 예상 소요: ${durMin}분`);
  if (r.memo) lines.push(`💬 고객 메모: ${r.memo}`);
  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push(`🆔 접수번호: ${r.as_id}`);
  if (r.created_at) lines.push(`📅 접수일시: ${formatDateKR(r.created_at)}`);
  if (r.status) {
    const statusKr =
      r.status === 'intake' ? '신규접수' :
      r.status === 'completed' ? '완료' :
      r.status === 'cancelled' ? '취소' :
      r.status;
    lines.push(`📋 상태: ${statusKr}`);
  }
  lines.push('━━━━━━━━━━━━━━━━━━━━━━');
  lines.push('');
  lines.push('🔗 TMS 상세보기:');
  lines.push(`${baseUrl}/repairs/${r.id}`);
  if (phone) {
    const phoneDigits = phone.replace(/\D/g, '');
    lines.push('');
    lines.push(`📞 바로 전화: tel:${phoneDigits}`);
  }
  lines.push('');
  lines.push('⚠ 이 일정은 MAMORU TMS에서 자동 생성됩니다.');
  lines.push('  변경·취소는 반드시 TMS에서 진행해 주세요.');
  const description = lines.join('\n');

  // 시작/종료 (visit_time 기준 + duration)
  const startTime = r.visit_time && r.visit_time.match(/^\d{1,2}:\d{2}/) ? r.visit_time.slice(0, 5) : '10:00';
  const endTime = addMinutes(startTime, durMin);
  const visitDate = r.visit_date || '';

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
    colorId: getRepairColorId(r.status),
    reminders: { useDefault: false, overrides: [] },
    extendedProperties: {
      private: {
        mamoru_repair_id: r.id,
        mamoru_repair_as_id: r.as_id,
        mamoru_repair_status: r.status,
        mamoru_source: 'repair_direct_visit',
        mamoru_version: '1.0',
      },
    },
  };

  return event;
}
