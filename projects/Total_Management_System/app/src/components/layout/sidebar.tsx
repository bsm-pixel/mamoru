'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';
import {
  LayoutDashboard,
  ShoppingCart,
  Settings,
  MessageSquare,
  Wrench,
  Package,
  Users,
  type LucideIcon,
} from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { NAV_ITEMS, NAV_ITEMS_FUTURE } from '@/lib/utils/constants';

const iconMap: Record<string, LucideIcon> = {
  LayoutDashboard,
  ShoppingCart,
  Settings,
  MessageSquare,
  Wrench,
  Package,
  Users,
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

      {/* 메인 내비 */}
      <nav className="flex-1 px-3 space-y-1">
        {NAV_ITEMS.map((item) => {
          const Icon = iconMap[item.icon];
          const active = pathname.startsWith(item.href);
          return (
            <Link
              key={item.href}
              href={item.href}
              className={cn(
                'flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition',
                active
                  ? 'bg-terracotta/15 text-terracotta'
                  : 'text-cream/60 hover:text-cream hover:bg-white/5'
              )}
            >
              {Icon && <Icon size={18} />}
              {item.label}
            </Link>
          );
        })}

        {/* 구분선 */}
        <div className="pt-4 pb-2">
          <div className="h-px bg-white/8" />
          <p className="text-[10px] text-cream/30 uppercase tracking-widest mt-3 mb-1 px-3">
            Coming Soon
          </p>
        </div>

        {NAV_ITEMS_FUTURE.map((item) => {
          const Icon = iconMap[item.icon];
          return (
            <span
              key={item.href}
              className="flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-cream/25 cursor-not-allowed"
            >
              {Icon && <Icon size={18} />}
              {item.label}
            </span>
          );
        })}
      </nav>
    </aside>
  );
}
