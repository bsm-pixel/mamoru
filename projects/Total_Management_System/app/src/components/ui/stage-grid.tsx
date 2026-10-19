'use client';

/**
 * 모바일 단계 버튼 판 — 가로 스크롤 탭 대신 한 화면에 모든 단계를 버튼으로 (2026-10-19)
 *
 * 쓰는 곳: 복원수리(8단계 4×2) · 판매관리(4상태) · 상담관리(종류·단계)
 * 규칙: 이름 작게 / 건수 크게 → "어디에 일이 쌓였나"가 1초에 보인다. 0건은 흐리게, 선택은 검정.
 *   count 가 null 이면 건수 대신 – (건수를 셀 수 없는 단계)
 */
export interface StageItem<K extends string> {
  key: K;
  label: string;
  count?: number | null;
  /** 건수 강조색 — 처리 필요처럼 손이 가야 하는 단계 */
  tone?: 'default' | 'amber';
}

interface Props<K extends string> {
  items: StageItem<K>[];
  active: K;
  onChange: (k: K) => void;
  cols?: 2 | 3 | 4;
  ariaLabel?: string;
}

const COLS = { 2: 'grid-cols-2', 3: 'grid-cols-3', 4: 'grid-cols-4' } as const;

export function StageGrid<K extends string>({ items, active, onChange, cols = 4, ariaLabel = '단계' }: Props<K>) {
  return (
    <div className={`grid ${COLS[cols]} gap-1.5`} role="tablist" aria-label={ariaLabel}>
      {items.map((it) => {
        const on = active === it.key;
        const hasCount = typeof it.count === 'number';
        const empty = hasCount && it.count === 0;
        return (
          <button
            key={it.key}
            type="button"
            role="tab"
            aria-selected={on}
            onClick={() => onChange(it.key)}
            className={`flex flex-col items-start justify-between min-h-[60px] px-2.5 py-2 rounded-xl border text-left transition active:scale-[0.98] ${
              on
                ? 'bg-stone-900 border-stone-900 text-white'
                : empty
                  ? 'bg-white/60 border-stone-200 text-stone-400'
                  : 'bg-white border-stone-200 text-stone-900 hover:border-stone-400'
            }`}
          >
            <span className={`text-[11px] font-semibold leading-tight whitespace-nowrap ${on ? 'text-white/80' : empty ? '' : 'text-stone-500'}`}>
              {it.label}
            </span>
            {hasCount ? (
              <span className={`text-lg font-bold tabular-nums leading-none mt-1 ${!on && !empty && it.tone === 'amber' ? 'text-amber-600' : ''}`}>{it.count}</span>
            ) : (
              <span className={`text-lg font-bold leading-none mt-1 ${on ? "text-white/50" : "text-stone-300"}`}>–</span>
            )}
          </button>
        );
      })}
    </div>
  );
}
