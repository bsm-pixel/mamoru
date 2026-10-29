/**
 * 경비 구분 3칸 — 고정(매달) · 연간 · 사용량 (2026-10-29, 마이그 167)
 *
 * DB는 kind(fixed|variable) + billing_cycle(annual|null) 두 칸. 화면은 한 줄 3칸 버튼으로 보여 준다.
 * 연간 = kind 'fixed' + billing_cycle 'annual' → 고정 지출 합계에 포함, 「월 환산 고정비」는 ÷12.
 * 경비 등록 폼(app/(dashboard)/expenses)과 메일 확인 대기(components/expenses/expense-inbox)가 같이 쓴다.
 */
export type ExpensePlan = 'fixed' | 'annual' | 'variable';

export const PLAN_OPTIONS: Array<{ value: ExpensePlan; short: string; long: string }> = [
  { value: 'fixed', short: '고정', long: '고정 (매달)' },
  { value: 'annual', short: '연간', long: '연간 (1년 1번)' },
  { value: 'variable', short: '사용량', long: '사용량 (쓴 만큼)' },
];

export function toPlan(kind: string | null | undefined, cycle: string | null | undefined): ExpensePlan {
  if (kind === 'fixed') return cycle === 'annual' ? 'annual' : 'fixed';
  return 'variable';
}

export function fromPlan(plan: ExpensePlan): { kind: 'fixed' | 'variable'; billing_cycle: 'annual' | null } {
  return { kind: plan === 'variable' ? 'variable' : 'fixed', billing_cycle: plan === 'annual' ? 'annual' : null };
}

/** 목록 꼬리표 — 미분류(옛 데이터)는 null */
export function planLabel(kind: string | null | undefined, cycle: string | null | undefined): string | null {
  if (!kind) return null;
  return PLAN_OPTIONS.find((o) => o.value === toPlan(kind, cycle))!.short;
}
