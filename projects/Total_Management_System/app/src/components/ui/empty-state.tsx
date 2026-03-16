import type { LucideIcon } from 'lucide-react';

interface EmptyStateProps {
  icon?: LucideIcon;
  message: string;
}

/** 목록 빈 상태 공통 컴포넌트 */
export function EmptyState({ icon: Icon, message }: EmptyStateProps) {
  return (
    <div className="flex flex-col items-center justify-center h-40 text-neutral-400">
      {Icon && <Icon size={32} className="mb-2 opacity-50" />}
      <span className="text-sm">{message}</span>
    </div>
  );
}
