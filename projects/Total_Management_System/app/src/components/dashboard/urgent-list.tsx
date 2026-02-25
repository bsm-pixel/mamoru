'use client';

import { cn } from '@/lib/utils/cn';
import { Badge } from '@/components/ui/badge';

interface UrgentItem {
  id: string;
  label: string;
  sublabel?: string;
  badge?: string;
  badgeColor?: string;
}

interface UrgentListProps {
  title: string;
  items: UrgentItem[];
  onItemClick?: (id: string) => void;
  emptyMessage?: string;
  maxItems?: number;
}

/** 긴급/처리 필요 건 리스트 (재사용 컴포넌트) */
export function UrgentList({
  title,
  items,
  onItemClick,
  emptyMessage = '처리 건 없음',
  maxItems = 5,
}: UrgentListProps) {
  const display = items.slice(0, maxItems);

  return (
    <div className="bg-card-white rounded-xl border border-neutral-200 shadow-sm overflow-hidden">
      <div className="px-4 py-3 border-b border-neutral-100">
        <h4 className="text-sm font-bold text-indigo-black">{title}</h4>
      </div>

      {display.length === 0 ? (
        <div className="px-4 py-6 text-center">
          <p className="text-sm text-neutral-400">{emptyMessage}</p>
        </div>
      ) : (
        <div className="divide-y divide-neutral-100">
          {display.map((item) => (
            <button
              key={item.id}
              type="button"
              onClick={() => onItemClick?.(item.id)}
              className={cn(
                'w-full flex items-center justify-between px-4 py-3 text-left transition',
                onItemClick ? 'hover:bg-warm-ivory/60 cursor-pointer' : ''
              )}
            >
              <div className="min-w-0">
                <p className="text-sm font-semibold text-indigo-black truncate">{item.label}</p>
                {item.sublabel && (
                  <p className="text-xs text-neutral-500 mt-0.5 truncate">{item.sublabel}</p>
                )}
              </div>
              {item.badge && (
                <Badge className={cn('shrink-0 ml-2', item.badgeColor || '')}>
                  {item.badge}
                </Badge>
              )}
            </button>
          ))}
        </div>
      )}

      {items.length > maxItems && (
        <div className="px-4 py-2 border-t border-neutral-100 text-center">
          <p className="text-xs text-neutral-400">+{items.length - maxItems}건 더</p>
        </div>
      )}
    </div>
  );
}
