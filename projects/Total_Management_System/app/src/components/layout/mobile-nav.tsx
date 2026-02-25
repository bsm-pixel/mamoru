'use client';

import Link from 'next/link';
import { usePathname } from 'next/navigation';
import { LayoutDashboard, ShoppingCart, MessageSquare, Settings, Wrench, type LucideIcon } from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { NAV_ITEMS } from '@/lib/utils/constants';

const iconMap: Record<string, LucideIcon> = {
  LayoutDashboard,
  ShoppingCart,
  MessageSquare,
  Wrench,
  Settings,
};

export function MobileNav() {
  const pathname = usePathname();

  return (
    <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-card-white/95 backdrop-blur-sm border-t border-neutral-200 safe-area-bottom">
      <div className="flex items-center justify-around h-14">
        {NAV_ITEMS.map((item) => {
          const Icon = iconMap[item.icon];
          const active = pathname.startsWith(item.matchPrefix);
          return (
            <Link
              key={item.href}
              href={item.href}
              className={cn(
                'flex flex-col items-center gap-0.5 min-w-[64px] py-1 transition',
                active ? 'text-terracotta' : 'text-neutral-400'
              )}
            >
              {Icon && <Icon size={20} />}
              <span className="text-[10px] font-medium">{item.label}</span>
            </Link>
          );
        })}
      </div>
    </nav>
  );
}
