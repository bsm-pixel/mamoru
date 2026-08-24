import type { ActivityTypes } from '@/hooks/use-activity-types';

/** 활동유형 칩 — 색은 고객 상세 타임라인과 통일 (주문=emerald, 구매=blue, 수리=rose) */
const CHIPS: { key: keyof ActivityTypes; label: string; cls: string }[] = [
  { key: 'has_order', label: '주문', cls: 'bg-emerald-50 text-emerald-700' },
  { key: 'has_sale', label: '구매', cls: 'bg-blue-50 text-blue-700' },
  { key: 'has_repair', label: '수리', cls: 'bg-rose-50 text-rose-700' },
];

/**
 * 고객 활동유형 칩 (주문/구매/수리) — 목록 이름 옆에 붙이는 초소형 배지.
 * types 가 없으면(로딩/활동없음) 아무것도 안 그림.
 */
export function ActivityChips({ types, className = '' }: { types?: ActivityTypes; className?: string }) {
  if (!types) return null;
  const active = CHIPS.filter((c) => types[c.key]);
  if (active.length === 0) return null;
  return (
    <span className={`inline-flex items-center gap-0.5 align-middle ${className}`}>
      {active.map((c) => (
        <span
          key={c.key}
          className={`inline-block px-1 py-px rounded text-[9px] font-semibold leading-none ${c.cls}`}
        >
          {c.label}
        </span>
      ))}
    </span>
  );
}
