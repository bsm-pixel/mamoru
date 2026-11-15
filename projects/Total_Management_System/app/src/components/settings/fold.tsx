import type { ReactNode } from 'react';
import { ChevronDown } from 'lucide-react';

/**
 * 설정 화면의 접는 칸 (2026-11-15)
 * 알림·연동 탭이 PC 3,500px · 모바일 4,600px 로 길어 한 번에 다 펼쳐 두지 않는다.
 * 제목 한 줄만 보이고 누르면 펼쳐진다. 자주 쓰는 것만 defaultOpen.
 */
export function Fold({ title, desc, defaultOpen = false, children }: { title: string; desc?: string; defaultOpen?: boolean; children: ReactNode }) {
  return (
    <details open={defaultOpen} className="group rounded-lg border border-neutral-200 bg-white">
      <summary className="flex items-center justify-between gap-3 px-4 py-3 cursor-pointer select-none list-none [&::-webkit-details-marker]:hidden">
        <span className="min-w-0">
          <span className="block text-sm font-bold text-neutral-800">{title}</span>
          {desc && <span className="block text-xs text-neutral-400 mt-0.5">{desc}</span>}
        </span>
        <ChevronDown size={16} className="text-neutral-400 shrink-0 transition-transform group-open:rotate-180" />
      </summary>
      <div className="px-4 pb-4 pt-1 space-y-4 border-t border-neutral-100">{children}</div>
    </details>
  );
}
