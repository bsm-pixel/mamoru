import { daysSinceKST } from '@/lib/utils/kst';

/**
 * 며칠째인지 작게 보여 주기 (2026-11-15) — "접수 3일째"
 * 오늘 들어온 건은 표시하지 않는다(1일째는 정보가 없다). warnDays 를 넘기면 굵게.
 *   <DaysSince date={r.received_at} warnDays={staleDays} />
 */
export function DaysSince({ date, warnDays, prefix = '', className = '' }: {
  date: string | null | undefined;
  warnDays?: number;
  prefix?: string;
  className?: string;
}) {
  if (!date) return null;
  const n = daysSinceKST(date);
  if (!Number.isFinite(n) || n < 1) return null;
  const late = typeof warnDays === 'number' && warnDays > 0 && n >= warnDays;
  return (
    <span className={`text-[11px] tabular-nums whitespace-nowrap ${late ? 'font-bold text-neutral-900' : 'text-neutral-400'} ${className}`}>
      {prefix}{n + 1}일째
    </span>
  );
}
