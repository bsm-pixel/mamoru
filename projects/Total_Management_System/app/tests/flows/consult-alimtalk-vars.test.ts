import { describe, it, expect, vi } from 'vitest';
import { ctx, jsonReq, idParam, settle } from '../helpers/setup';

/**
 * 상담 알림톡 — Make 분기가 읽는 값을 TMS가 실제로 보내는가 (2026-11-22)
 *
 * 실제 사고: 출장 접수 알림톡의 「희망날짜 · 희망시간대」가 빈칸으로 나갔다.
 *   Make 「01 상담」 분기는 템플릿 칸을 아래 이름에서 읽는데, TMS가 다른 이름(date · time)으로만 보내고 있었다.
 *   (출장 접수는 Make 쪽 연결을 hope_days_text · hope_times_text 로 고쳤다.)
 * 아래 표는 Make 시나리오 9837024 의 분기별 연결을 옮긴 것이다. Make 를 바꾸면 이 표도 같이 고친다.
 */
const MAKE_READS: Record<string, string[]> = {
  confirmed: ['name', 'visit_date', 'visit_time', 'change_request_link'],
  rescheduled: ['new_date', 'new_time', 'change_request_link'],
  request: ['name', 'address', 'hope_days_text', 'hope_times_text'],
  field_confirmed: ['address', 'visit_date', 'visit_time', 'change_request_link'],
  field_rescheduled: ['address', 'visit_date', 'visit_time', 'change_request_link'],
  field_cancelled: ['address', 'visit_date', 'visit_time'],
};

/** 보내려던 알림톡 한 통에서 Make 가 읽는 칸이 모두 채워졌는지 */
function missing(p: { template: string; name?: string; data?: Record<string, unknown> }) {
  const body: Record<string, unknown> = { name: p.name, ...(p.data || {}) };
  return (MAKE_READS[p.template] || []).filter((k) => body[k] === undefined || body[k] === null || String(body[k]).trim() === '');
}

// 접수 처리 안에서 부르는 바깥 기능 — 메일 · 달력 · 좌표 · 고객 매칭은 이 점검과 무관하다
vi.mock('@/lib/notification/email', () => ({ sendAdminEmail: vi.fn(async () => true) }));
vi.mock('@/lib/google/calendar-sync', () => ({ syncConsultationToCalendar: vi.fn(async () => {}) }));
vi.mock('@/lib/kakao/geocode', () => ({ geocodeForConsultation: vi.fn(async () => null) }));
vi.mock('@/lib/customer/match-or-create', () => ({ matchOrCreateCustomer: vi.fn(async () => ({ customerId: 'c1' })) }));

describe('상담 알림톡 — Make 가 읽는 칸이 빈칸으로 나가지 않는다', () => {
  it('출장 접수: 희망 요일 · 선호 시간대가 실린다', async () => {
    const { POST } = await import('@/app/api/consultation/public/submit/route');
    const res = await POST(jsonReq('/api/consultation/public/submit', 'POST', {
      name: '점검고객', phone: '010-0000-0000', type: '출장 요청',
      addressRoad: '인천 남동구 점검로 1', addressDetail: '2층', days: ['월', '수'], timePrefs: ['오전'],
    }));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    await settle();
    const p = ctx.sent.find((s) => s.template === 'request');
    expect(p, '출장 접수 알림톡을 보내려 해야 한다').toBeTruthy();
    expect(missing(p)).toEqual([]);
    expect(p.data.hope_days_text).toBe('월, 수');
    expect(p.data.hope_times_text).toBe('오전');
  });

  it('출장 접수에서 요일 · 시간대를 고르지 않았거나 「무관」이면: 빈칸 대신 「상관없음」', async () => {
    const { POST } = await import('@/app/api/consultation/public/submit/route');
    const res = await POST(jsonReq('/api/consultation/public/submit', 'POST', {
      name: '점검고객', phone: '010-0000-0003', type: '출장 요청',
      addressRoad: '인천 남동구 점검로 1', days: [], timePrefs: ['무관'],
    }));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    await settle();
    const p = ctx.sent.find((s) => s.template === 'request');
    expect(missing(p)).toEqual([]);
    expect([p.data.hope_days_text, p.data.hope_times_text]).toEqual(['상관없음', '상관없음']);
  });

  it('매장방문 즉시 확정: 방문 날짜 · 시간이 실린다', async () => {
    const { POST } = await import('@/app/api/consultation/public/submit/route');
    const res = await POST(jsonReq('/api/consultation/public/submit', 'POST', {
      name: '점검고객', phone: '010-0000-0001', type: '매장 방문', visitDate: '2026-12-01', visitTime: '14:00',
    }));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    await settle();
    const p = ctx.sent.find((s) => s.template === 'confirmed');
    expect(p, '매장방문 확정 알림톡을 보내려 해야 한다').toBeTruthy();
    expect(missing(p)).toEqual([]);
    expect([p.data.visit_date, p.data.visit_time]).toEqual(['2026-12-01', '14:00']);
  });

  it('관리자가 매장방문 일정을 바꾸면: 바뀐 날짜 · 시간이 실린다', async () => {
    ctx.db.seed('consultations', [{
      id: 'k1', unique_id: 'u1', name: '점검고객', phone: '01000000002', consultation_type: 'store_visit',
      status: 'confirmed', visit_date: '2026-12-01', visit_time: '14:00', address_road: null, address_detail: null, memo: null,
    }]);
    const { PATCH } = await import('@/app/api/consultation/[id]/route');
    const res = await PATCH(jsonReq('/api/consultation/k1', 'PATCH', { visit_date: '2026-12-03', visit_time: '11:00' }), idParam('k1'));
    expect(res.status, JSON.stringify(await res.clone().json())).toBe(200);
    await settle();
    const p = ctx.sent.find((s) => s.template === 'rescheduled');
    expect(p, '일정변경 알림톡을 보내려 해야 한다').toBeTruthy();
    expect(missing(p)).toEqual([]);
    expect([p.data.new_date, p.data.new_time]).toEqual(['2026-12-03', '11:00']);
  });
});
