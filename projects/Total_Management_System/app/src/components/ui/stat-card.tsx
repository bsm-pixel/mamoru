'use client';

/**
 * StatCard — TMS 공통 통계 카드 (시안 B 톤, 2026-05-27 사장님 채택)
 *
 * 사용처:
 *   - /dashboard 4카드 (주문/상담/수리/판매)
 *   - /consultations 요약 3카드 (신규/진행/완료)
 *   - /repairs 상태 4카드 (신규/진행/미입금/3일경과)
 *   - /orders/dashboard 4카드 (결제완료/준비중/배송중/오늘주문)
 *   - 향후 매입/시리얼/계약서 등 모든 통계 카드
 *
 * 디자인 룰 (memory/feedback_tms_design_direction.md):
 *   - rounded-2xl + border border-stone-200 + bg-white
 *   - 헤더: 작은 아이콘 + UPPERCASE 라벨 + 우측 화살표
 *   - 본문: 큰 숫자(3xl) + accent 색 + 작은 보조 라벨
 *   - 활성/필터: ring-1 ring-stone-900 + border-stone-900
 *   - hover: border-stone-300 + 화살표 translate
 *
 * Props 3분류 (feedback_code_dry_no_duplicates.md):
 *   - 데이터: label, value, subLabel, icon
 *   - 스타일 변형: accent, active
 *   - 콜백: href | onClick (하나만 — href 우선)
 */

import Link from 'next/link';
import { ArrowRight } from 'lucide-react';
import type { LucideIcon } from 'lucide-react';

type Accent = 'blue' | 'amber' | 'emerald' | 'rose' | 'orange' | 'violet' | 'stone';

const ACCENT_TEXT: Record<Accent, string> = {
  blue:    'text-blue-600',
  amber:   'text-amber-600',
  emerald: 'text-emerald-600',
  rose:    'text-rose-600',
  orange:  'text-orange-600',
  violet:  'text-violet-600',
  stone:   'text-stone-800',
};

export interface StatCardProps {
  label: string;
  value: number | string;
  icon: LucideIcon;
  /** 큰 숫자 바로 아래 (text-[10px] stone-500) — 예: "결제완료" / "확인 필요" */
  primarySub?: string;
  /** 그 아래 border-top + truncate (text-[10px] stone-400) — 예: "준비 3 · 완료 8" */
  secondarySub?: string;
  accent?: Accent;
  /** 0이면 회색으로 강제 (사용 안 함 / 알림 0건 등) */
  dimWhenZero?: boolean;
  /** 활성/필터 선택 상태 (ring 표시) */
  active?: boolean;
  /** 컴팩트 모드 — 밀집 배치용(작은 패딩·숫자, 화살표 생략). 예: /repairs 좌측 통계 */
  compact?: boolean;
  /** 클릭 시 라우팅 (href 우선) */
  href?: string;
  onClick?: () => void;
}

export function StatCard({
  label, value, icon: Icon, primarySub, secondarySub,
  accent = 'stone', dimWhenZero = false, active = false, compact = false,
  href, onClick,
}: StatCardProps) {
  const isZero = typeof value === 'number' && value === 0;
  const accentClass = (dimWhenZero && isZero) ? 'text-stone-300' : ACCENT_TEXT[accent];

  const className = `bg-white rounded-2xl border hover:border-stone-300 transition group text-left w-full h-full flex flex-col ${
    compact ? 'p-2.5' : 'p-4'
  } ${active ? 'border-stone-900 ring-1 ring-stone-900' : 'border-stone-200'}`;

  const content = (
    <>
      <div className={`flex items-center justify-between ${compact ? 'mb-1' : 'mb-2'}`}>
        <div className="flex items-center gap-1.5 min-w-0">
          <Icon size={compact ? 12 : 13} className="text-stone-400 shrink-0" />
          <p className={`${compact ? 'text-[10px]' : 'text-[11px]'} text-stone-500 uppercase tracking-wider font-semibold truncate`}>{label}</p>
        </div>
        {!compact && <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />}
      </div>
      <div className="flex-1 flex flex-col justify-center">
        <p className={`${compact ? 'text-xl' : 'text-3xl'} font-bold leading-none ${accentClass}`}>{value}</p>
        {primarySub && <p className={`text-[10px] text-stone-500 ${compact ? 'mt-0.5' : 'mt-1'}`}>{primarySub}</p>}
      </div>
      {secondarySub && (
        <p className="text-[10px] text-stone-400 mt-2 pt-2 border-t border-stone-100 truncate">{secondarySub}</p>
      )}
    </>
  );

  if (href) {
    return <Link href={href} className={className}>{content}</Link>;
  }
  return (
    <button onClick={onClick} className={className} type="button">
      {content}
    </button>
  );
}
