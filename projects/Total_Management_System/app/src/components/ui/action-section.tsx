'use client';

import { useState, type ReactNode } from 'react';
import { ChevronDown } from 'lucide-react';

/**
 * 상세 패널 「액션」 영역 표준 구성요소 (2026-09-15)
 *
 * 문제: 액션 카드가 버튼 나열이 되면서 위계가 뒤집혔다.
 *   납품 상세에서 가장 큰 검은 버튼이 제일 드물게 쓰는 [송장 재발급] 이었고,
 *   주 행동인 [출고 완료] 는 흰 버튼으로 박스 안에 갇혀 있었으며,
 *   출고가 끝난 뒤에도 [송장 취소] 만 빨갛게 덩그러니 남았다. (사장님 지적)
 *
 * 규칙 — 액션 영역은 항상 이 3단으로만 쌓는다:
 *   1) ActionNote   지금 상태 한 줄 (읽기 전용, 버튼 아님)
 *   2) 주 액션       Button 1개. "지금 눌러야 할 것" 하나만 검은 버튼
 *   3) MoreActions  드물게 쓰는 것·되돌리기 → 접어둔다 (평소엔 글자 한 줄)
 *   4) DangerZone   파괴적 액션. hairline 아래 작은 회색, hover 에서만 빨강
 *
 * MAMORU 톤: 모노크롬 · 장식 금지 · 여백이 위계. 빨강은 hover 에만.
 */

type Tone = 'wait' | 'done' | 'muted';

const NOTE_TONE: Record<Tone, string> = {
  wait: 'bg-amber-50 text-amber-700',
  done: 'bg-green-50 text-green-700',
  muted: 'bg-stone-50 text-neutral-600',
};

interface NoteProps {
  tone?: Tone;
  /** 굵게 나가는 상태명 — 예: 출고대기 */
  title: string;
  /** 상태명 옆 보조 정보 — 예: 롯데택배 317653777442 */
  meta?: ReactNode;
  /** 왜 기다리는지·다음에 뭐가 자동으로 되는지 */
  children?: ReactNode;
}

/** 지금 상태 한 줄. 버튼처럼 보이면 안 된다 */
export function ActionNote({ tone = 'muted', title, meta, children }: NoteProps) {
  return (
    <div className={`rounded-lg px-2.5 py-2 ${NOTE_TONE[tone]}`}>
      <p className="text-xs font-semibold leading-relaxed">
        {title}
        {meta && <span className="ml-1.5 font-normal opacity-80">{meta}</span>}
      </p>
      {children && <p className="text-[11px] leading-relaxed opacity-80 mt-0.5">{children}</p>}
    </div>
  );
}

interface MoreProps {
  /** 접힘 상태에서 보이는 한 줄 — 질문형으로 쓴다 ("다르게 보냈어요", "잘못 처리했나요?") */
  label: string;
  defaultOpen?: boolean;
  children: ReactNode;
}

/** 드물게 쓰는 액션 묶음 — 평소엔 글자 한 줄로 접혀 있어 주 액션을 가리지 않는다 */
export function MoreActions({ label, defaultOpen = false, children }: MoreProps) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div>
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        className="w-full flex items-center justify-center gap-1 py-1.5 text-xs text-neutral-500 hover:text-neutral-800 transition"
      >
        {label}
        <ChevronDown size={13} className={`transition-transform ${open ? 'rotate-180' : ''}`} />
      </button>
      {open && <div className="space-y-1.5 pt-0.5">{children}</div>}
    </div>
  );
}

/** 파괴적 액션 구역 — 주 액션과 시각적으로 분리하고 무게를 낮춘다 */
export function DangerZone({ children }: { children: ReactNode }) {
  return <div className="pt-2 mt-1 border-t border-neutral-100 space-y-0.5">{children}</div>;
}

/** 위험/되돌리기용 텍스트 버튼 — 평소 회색, hover 에서만 빨강 */
export function DangerLink({
  onClick,
  disabled,
  children,
}: {
  onClick: () => void;
  disabled?: boolean;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className="w-full text-center text-xs text-neutral-400 hover:text-red-600 py-1 transition disabled:opacity-50"
    >
      {children}
    </button>
  );
}

/** 접힘 안쪽에서 쓰는 보조 액션(파괴적이지 않은 것) — 예: 송장 재발급 */
export function SubtleButton({
  onClick,
  disabled,
  children,
}: {
  onClick: () => void;
  disabled?: boolean;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      className="w-full h-9 rounded-lg border border-neutral-200 text-xs text-neutral-600 hover:bg-neutral-50 transition disabled:opacity-50"
    >
      {children}
    </button>
  );
}
