'use client';

/**
 * RevenueDarkCard — 매출/돈 합계용 어두운 카드 (A2 그라데이션 톤, 2026-05-27 사장님 채택)
 *
 * 사용처:
 *   - /repairs 상단 매출 KPI (이번달 복원수리)
 *   - /sales B2C/B2B 매출 카드 (안 3 채택안)
 *   - 향후 /purchasing 매입 합계 등
 *
 * 룰: "매출/돈 합계 = 어두운 카드 (부드러운 톤)"
 *
 * 디자인:
 *   - bg-gradient-to-br from-stone-800 to-stone-900 + ring-1 ring-white/5 (부담 없는 깊이감)
 *   - rounded-2xl
 *   - 헤더: 아이콘 + UPPERCASE 라벨 + 큰 숫자 + 우측 보조 (예: 자루 수)
 *   - 하단 분할 셀: bg-stone-900/40 (옅은 분리)
 *
 * Props:
 *   - 헤더: icon, label, amount, rightValue?
 *   - 분할 셀: splits[] (선택)
 *   - 하단 그리드: bottomGrid[] (선택, 다른 형식 — sales 페이지 처럼 이번주/미수금 같이)
 */

import type { LucideIcon } from 'lucide-react';

export interface RevenueSplit {
  label: string;
  amount: string;   // 이미 포맷된 ₩ 문자열
  sub?: string;     // 예: "5자루"
}

export interface RevenueBottomGrid {
  label: string;
  value: string;       // 이미 포맷된 문자열
  /** 강조 색 (예: 미수금 amber-300) */
  highlight?: 'amber' | 'rose';
}

export interface RevenueDarkCardProps {
  /** 좌측 헤더 아이콘 (TrendingUp 등) — 생략 시 아이콘 없이 라벨부터 시작 */
  icon?: LucideIcon;
  label: string;
  amount: string;         // 이미 포맷된 ₩ 문자열
  /** amount 아래 부가 텍스트 (예: "이번달 · 5건") */
  amountSub?: string;
  /** 우측 상단 보조 값 (예: "32자루") */
  rightValue?: string;
  rightValueSub?: string; // 예: "자루"
  /** 하단 3분할 셀 (예: 마모루/타사/B2B) */
  splits?: RevenueSplit[];
  /** 하단 2분할 그리드 (예: 이번주/미수금) — splits 와 동시 사용 X, 둘 중 하나만 */
  bottomGrid?: RevenueBottomGrid[];
}

const HIGHLIGHT_TEXT = {
  amber: 'text-amber-300',
  rose:  'text-rose-300',
} as const;

export function RevenueDarkCard({
  icon: Icon, label, amount, amountSub, rightValue, rightValueSub = '자루',
  splits, bottomGrid,
}: RevenueDarkCardProps) {
  return (
    <div className="rounded-2xl bg-gradient-to-br from-stone-800 to-stone-900 text-white overflow-hidden ring-1 ring-white/5">
      {/* 헤더 */}
      <div className="flex items-center gap-3 px-5 py-4">
        {Icon && <Icon size={18} className="opacity-60" />}
        <div className="min-w-0">
          <p className="text-[11px] uppercase tracking-wider opacity-60 font-semibold">{label}</p>
          <p className="text-2xl font-bold">{amount}</p>
          {amountSub && <p className="text-xs opacity-70 mt-0.5">{amountSub}</p>}
        </div>
        {rightValue && (
          <p className="ml-auto text-xl font-bold">
            {rightValue}<span className="text-xs opacity-60 ml-0.5">{rightValueSub}</span>
          </p>
        )}
      </div>

      {/* 하단 — splits 우선 */}
      {splits && splits.length > 0 && (
        <div className="grid grid-cols-3 gap-px bg-white/10">
          {splits.map((s) => (
            <div key={s.label} className="px-3 py-2.5 text-center bg-stone-900/40">
              <p className="text-[10px] opacity-50 uppercase tracking-wider">{s.label}</p>
              <p className="text-sm font-bold mt-0.5">{s.amount}</p>
              {s.sub && <p className="text-[10px] opacity-40">{s.sub}</p>}
            </div>
          ))}
        </div>
      )}

      {bottomGrid && bottomGrid.length > 0 && !splits && (
        <div className="px-5 pb-4 pt-3 border-t border-white/10 grid grid-cols-2 gap-2 text-xs">
          {bottomGrid.map((g) => {
            const cls = g.highlight ? HIGHLIGHT_TEXT[g.highlight] : 'opacity-90';
            return (
              <div key={g.label}>
                <span className={`${g.highlight ? cls : 'opacity-60'}`}>{g.label}</span>
                <div className={`font-semibold mt-0.5 ${cls}`}>{g.value}</div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
