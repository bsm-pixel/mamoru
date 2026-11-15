/**
 * 상태 색 표 (2026-11-15 통일) — 상태 꼬리표의 색은 이 다섯 가지뿐
 *
 *   대기 wait      노랑  내 손이 필요하거나 상대의 답을 기다리는 중
 *   진행 progress  파랑  일이 굴러가는 중 (예약됨 · 수리 중 · 배송 중)
 *   완료 done      초록  끝남
 *   취소 cancel    회색  취소 · 작성 중 · 보류 — 더 할 일이 없음
 *   문제 problem   빨강  반품 요청 · 미결제 · 불량 — 바로 봐야 함
 *
 * 전엔 화면 12곳이 저마다 색 표를 들고 있어 같은 「완료」가 초록·에메랄드·회색으로, 「취소」가 빨강으로 보였다.
 * 새 상태를 만들면 색을 고르지 말고 이 다섯 중 무엇인지만 정한다.
 * (일정 화면의 매장/출장/수리/수거 색은 "종류" 색이고 구글 캘린더 색과 맞춘 것이라 여기 해당하지 않는다 — lib/schedule/colors.ts)
 */
export type StatusTone = 'wait' | 'progress' | 'done' | 'cancel' | 'problem';

export const TONE: Record<StatusTone, string> = {
  wait: 'bg-warning-soft text-warning',
  progress: 'bg-info-soft text-info',
  done: 'bg-success-soft text-success',
  cancel: 'bg-neutral-100 text-neutral-500',
  problem: 'bg-error-soft text-error',
};

export const TONE_LABEL: Record<StatusTone, string> = {
  wait: '대기', progress: '진행', done: '완료', cancel: '취소', problem: '문제',
};

/** 상태 → 다섯 가지 중 하나를 적은 표를 색 글자 표로 바꾼다 */
export function toneClasses<K extends string>(map: Record<K, StatusTone>): Record<K, string> {
  const out = {} as Record<K, string>;
  (Object.keys(map) as K[]).forEach((k) => { out[k] = TONE[map[k]]; });
  return out;
}
