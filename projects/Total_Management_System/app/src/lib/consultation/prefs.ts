/**
 * 출장 요청 — 고객이 고른 희망 요일 · 선호 시간대를 보여 주는 말 (2026-11-22 사장님 결정)
 *
 * 고르지 않았으면 「상관없음」(= 아무 때나 괜찮다). 시간대의 「무관」 버튼도 같은 뜻이라 「상관없음」으로 보인다.
 * 전엔 고르지 않으면 알림톡에 빈칸으로 나갔고, TMS 화면에는 아무것도 안 보였다.
 * 저장된 값(gas_raw.days · gas_raw.timePrefs)은 고객이 보낸 그대로 두고, 보여 줄 때만 이 함수를 거친다.
 * 알림톡(api/consultation/public/submit) · 상담 상세 · 출장 목록 · 지도 · 시간 제안 창이 같이 쓴다.
 */
export const ANY_PREF = '상관없음';

const split = (v: unknown): string[] =>
  (Array.isArray(v) ? v.map(String) : String(v ?? '').split(',')).map((s) => s.trim()).filter(Boolean);

/** 희망 요일 목록 — 비었으면 ['상관없음'] */
export function prefDaysOf(v: unknown): string[] {
  const days = split(v);
  return days.length > 0 ? days : [ANY_PREF];
}

/** 선호 시간대 목록 — 비었거나 「무관」을 골랐으면 ['상관없음'] */
export function prefTimesOf(v: unknown): string[] {
  const times = split(v);
  return times.length === 0 || times.includes('무관') ? [ANY_PREF] : times;
}

/** 알림톡에 넣는 한 줄 — 「월, 수」 / 「상관없음」 */
export const prefText = (list: string[]) => list.join(', ');
