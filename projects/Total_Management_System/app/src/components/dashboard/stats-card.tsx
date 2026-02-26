import { cn } from '@/lib/utils/cn';
import type { LucideIcon } from 'lucide-react';

interface StatsCardProps {
  label: string;
  value: string | number;
  icon: LucideIcon;
  color?: string;
  bgColor?: string;
  subtitle?: string;  // R2: 부가 설명
}

export function StatsCard({
  label,
  value,
  icon: Icon,
  color = 'text-terracotta',
  bgColor = 'bg-terracotta/10',
  subtitle,
}: StatsCardProps) {
  return (
    <div className="bg-card-white rounded-xl border border-neutral-200 p-4 shadow-sm">
      <div className="flex items-center gap-3">
        <div className={cn('w-10 h-10 rounded-lg flex items-center justify-center', bgColor)}>
          <Icon size={20} className={color} />
        </div>
        <div>
          <p className="text-xs text-neutral-500">{label}</p>
          <p className="text-xl font-bold text-indigo-black mt-0.5">{value}</p>
          {subtitle && <p className="text-[10px] text-neutral-400 mt-0.5">{subtitle}</p>}
        </div>
      </div>
    </div>
  );
}
