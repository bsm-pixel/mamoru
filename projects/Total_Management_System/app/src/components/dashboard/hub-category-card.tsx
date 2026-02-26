import Link from 'next/link';
import { cn } from '@/lib/utils/cn';
import type { LucideIcon } from 'lucide-react';

interface StatItem {
  label: string;
  value: number;
  color?: string; // text-warning, text-info 등
}

interface HubCategoryCardProps {
  title: string;
  icon: LucideIcon;
  href: string;
  stats: StatItem[];
  summary?: string;  // R3: 하단 요약 라인 (예: "이번주 ₩420,000")
}

export function HubCategoryCard({ title, icon: Icon, href, stats, summary }: HubCategoryCardProps) {
  return (
    <Link
      href={href}
      className="group block bg-card-white rounded-xl border border-neutral-200 shadow-sm hover:shadow-md transition-shadow overflow-hidden"
    >
      <div className="flex">
        {/* 좌측 accent border */}
        <div className="w-1 bg-terracotta shrink-0" />

        <div className="flex-1 p-4 md:p-5">
          {/* 헤더 */}
          <div className="flex items-center gap-2 mb-3">
            <Icon size={20} className="text-terracotta" />
            <h3 className="text-sm font-bold text-indigo-black">{title}</h3>
          </div>

          {/* 수치 */}
          <div className="flex items-center gap-4">
            {stats.map((s) => (
              <div key={s.label} className="flex-1 min-w-0">
                <p className="text-xs text-neutral-500 truncate">{s.label}</p>
                <p className={cn('text-2xl font-bold mt-0.5', s.color || 'text-indigo-black')}>
                  {s.value}
                  <span className="text-xs font-normal text-neutral-400 ml-0.5">건</span>
                </p>
              </div>
            ))}
          </div>

          {/* R3: 요약 라인 */}
          {summary && (
            <p className="mt-2 pt-2 border-t border-neutral-100 text-xs text-neutral-500">
              {summary}
            </p>
          )}
        </div>
      </div>
    </Link>
  );
}
