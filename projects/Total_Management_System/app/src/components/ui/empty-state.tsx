import type { LucideIcon } from 'lucide-react';

interface EmptyStateProps {
  icon?: LucideIcon;
  /** 한 문장 — "발주 내역이 없습니다" */
  message: string;
  /** 다음에 할 일 버튼 1개 (있을 때만) */
  action?: { label: string; onClick: () => void };
  /** 표 안 · 좁은 칸용 — 높이를 줄이고 아이콘을 뺀다 */
  compact?: boolean;
}

/**
 * 빈 화면 공통 틀 (2026-11-15 통일) — 한 문장 + 버튼 1개
 * 전엔 화면마다 문구 · 크기가 제각각이고 "그래서 무엇을 하면 되는지"가 없었다.
 */
export function EmptyState({ icon: Icon, message, action, compact }: EmptyStateProps) {
  return (
    <div className={`flex flex-col items-center justify-center text-neutral-400 ${compact ? 'py-8 px-4' : 'min-h-40 py-6 px-4'}`}>
      {Icon && !compact && <Icon size={32} className="mb-2 opacity-50" />}
      <span className="text-sm text-center">{message}</span>
      {action && (
        <button
          type="button"
          onClick={action.onClick}
          className="mt-3 h-9 px-4 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition"
        >
          {action.label}
        </button>
      )}
    </div>
  );
}
