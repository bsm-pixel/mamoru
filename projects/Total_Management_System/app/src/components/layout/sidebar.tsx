'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';
import {
  LayoutDashboard,
  ShoppingCart,
  Settings,
  MessageSquare,
  Wrench,
  Store,
  Package,
  Users,
  Truck,
  Boxes,
  BarChart3,
  Star,
  Building2,
  type LucideIcon,
} from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { NAV_GROUPS } from '@/lib/utils/constants';

const iconMap: Record<string, LucideIcon> = {
  LayoutDashboard,
  ShoppingCart,
  Settings,
  MessageSquare,
  Wrench,
  Store,
  Package,
  Users,
  Truck,
  Boxes,
  BarChart3,
  Star,
  Building2,
};

export function Sidebar() {
  const pathname = usePathname();

  return (
    <aside className="hidden md:flex flex-col w-56 min-h-screen bg-indigo-black text-cream border-r border-white/8">
      {/* 로고 */}
      <div className="px-5 py-5">
        <h1 className="text-lg font-extrabold tracking-tight">MAMORU</h1>
        <p className="text-[11px] text-cream/50 mt-0.5">TMS v1.0</p>
      </div>

      {/* 그룹별 내비 */}
      <nav className="flex-1 px-3 pb-4 overflow-y-auto">
        {NAV_GROUPS.map((group, gi) => (
          <div key={gi} className={gi > 0 ? 'mt-4' : ''}>
            {/* 그룹 레이블 */}
            {group.group && (
              <p className="px-3 mb-1.5 text-[10px] font-semibold tracking-wider uppercase text-cream/30">
                {group.group}
              </p>
            )}

            {/* 메뉴 항목 */}
            <div className="space-y-0.5">
              {group.items.map((item) => {
                const Icon = iconMap[item.icon];
                const active = pathname.startsWith(item.matchPrefix);
                return (
                  <Link
                    key={item.href}
                    href={item.href}
                    className={cn(
                      'flex items-center gap-3 px-3 py-2 rounded-lg text-sm font-medium transition',
                      active
                        ? 'bg-white/10 text-cream border-l-2 border-cream'
                        : 'text-cream/40 hover:text-cream/80 hover:bg-white/5'
                    )}
                  >
                    {Icon && <Icon size={18} />}
                    {item.label}
                  </Link>
                );
              })}
            </div>
          </div>
        ))}
      </nav>
    </aside>
  );
}
