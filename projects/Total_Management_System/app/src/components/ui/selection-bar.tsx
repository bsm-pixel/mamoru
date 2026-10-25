'use client';

import type { ReactNode } from 'react';
import { X } from 'lucide-react';

/**
 * 목록 다중선택 공통 (2026-10-25 IA 통일 — 복원수리·주문·판매 목록).
 * - PC: 목록 왼쪽 체크박스 칸을 항상 보여 주고, 1건 이상 체크하면 SelectionBar 가 뜬다(모드 전환 버튼 없음).
 * - 모바일: [선택] 으로 선택 모드에 들어간 뒤 행을 눌러 고른다.
 * - 버튼 이름은 상세 화면과 같게(준비표·꼬리표 …) — "1건은 상세에서, 여러 건은 골라서 막대에서".
 */

export interface SelectionAction {
  key: string;
  label: string;
  icon?: ReactNode;
  onClick: () => void;
  disabled?: boolean;
}

export function SelectionBar({ count, actions, onClear }: { count: number; actions: SelectionAction[]; onClear: () => void }) {
  if (count <= 0) return null;
  return (
    <div className="flex items-center gap-2 flex-wrap rounded-lg border border-neutral-900 bg-white px-3 py-2">
      <span className="text-xs font-bold text-neutral-900 tabular-nums">{count}건 선택됨</span>
      <div className="flex items-center gap-1.5 flex-wrap">
        {actions.map((a) => (
          <button
            key={a.key}
            type="button"
            onClick={a.onClick}
            disabled={a.disabled}
            className="flex items-center gap-1 text-xs font-semibold text-neutral-700 border border-neutral-200 rounded-md px-2.5 py-1.5 hover:bg-neutral-50 transition disabled:opacity-50"
          >
            {a.icon}
            {a.label}
          </button>
        ))}
      </div>
      <button type="button" onClick={onClear} className="ml-auto flex items-center gap-1 text-xs text-neutral-500 hover:text-neutral-800">
        <X size={13} /> 선택 해제
      </button>
    </div>
  );
}

/** 행 체크박스 — 행 클릭(상세 열기)과 분리. 체크박스만 눌러도 선택/해제 */
export function RowCheck({ checked, onToggle, className = '' }: { checked: boolean; onToggle: () => void; className?: string }) {
  return (
    <input
      type="checkbox"
      checked={checked}
      onChange={onToggle}
      onClick={(e) => e.stopPropagation()}
      aria-label="선택"
      className={`w-4 h-4 rounded border-neutral-300 accent-neutral-900 cursor-pointer shrink-0 ${className}`}
    />
  );
}
