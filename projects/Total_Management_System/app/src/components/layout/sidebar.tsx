'use client';

import { useState, useEffect } from 'react';
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
  Handshake,
  PackageOpen,
  PackageSearch,
  Search,
  FileSignature,
  Wallet,
  CalendarOff,
  Zap,
  Tag,
  ChevronDown,
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
  Handshake,
  PackageOpen,
  PackageSearch,
  Search,
  FileSignature,
  Wallet,
  CalendarOff,
  Zap,
  Tag,
};

const COLLAPSE_KEY = 'tms-sidebar-collapsed';

export function Sidebar() {
  const pathname = usePathname();
  // 그룹 접힘 상태 (localStorage 영속)
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  useEffect(() => {
    try {
      const s = localStorage.getItem(COLLAPSE_KEY);
      if (s) setCollapsed(JSON.parse(s));
    } catch { /* ignore */ }
  }, []);
  const toggleGroup = (key: string) => {
    setCollapsed((c) => {
      const next = { ...c, [key]: !c[key] };
      try { localStorage.setItem(COLLAPSE_KEY, JSON.stringify(next)); } catch { /* ignore */ }
      return next;
    });
  };

  // matchPrefix 정확 매칭 (하위 경로가 별도 메뉴인 경우 겹치지 않게 — IA 정리 2026-05-18)
  const isActive = (matchPrefix: string) =>
    matchPrefix === '/consultations'
      ? pathname === '/consultations' || (pathname.startsWith('/consultations/') && !pathname.startsWith('/consultations/calendar'))
      : pathname.startsWith(matchPrefix);

  const dashActive = pathname.startsWith('/dashboard');

  return (
    <aside className="hidden md:flex flex-col w-56 h-screen sticky top-0 self-start bg-indigo-black text-cream border-r border-white/8">
      {/* 로고 + 대시보드(홈) 아이콘 — 대시보드는 메뉴 목록에서 빼고 여기로 */}
      <div className="px-4 py-4 flex items-center justify-between flex-shrink-0">
        <div>
          <h1 className="text-lg font-extrabold tracking-tight">MAMORU</h1>
          <p className="text-[11px] text-cream/50 mt-0.5">TMS v1.0</p>
        </div>
        <Link
          href="/dashboard"
          title="대시보드 (홈)"
          aria-label="대시보드"
          className={cn(
            'w-9 h-9 rounded-lg flex items-center justify-center transition flex-shrink-0',
            dashActive ? 'bg-white/15 text-cream' : 'text-cream/55 hover:bg-white/8 hover:text-cream/90'
          )}
        >
          <LayoutDashboard size={18} />
        </Link>
      </div>

      {/* 그룹별 내비 (그룹 접기 + pill 액티브) */}
      <nav className="flex-1 px-2.5 pb-4 overflow-y-auto">
        {NAV_GROUPS.map((group, gi) => {
          const items = group.items.filter((it) => it.href !== '/dashboard'); // 대시보드 제외
          if (items.length === 0) return null; // 대시보드만 있던 빈 그룹 스킵
          const hasHeader = !!group.group;
          const isCollapsed = hasHeader && !!collapsed[group.group];
          return (
            <div key={gi}>
              {hasHeader && (
                <button
                  type="button"
                  onClick={() => toggleGroup(group.group)}
                  className="w-full flex items-center justify-between px-2.5 mt-3 mb-1 text-cream/45 hover:text-cream/75 transition"
                >
                  <span className="text-[11px] font-semibold tracking-wider uppercase">{group.group}</span>
                  <ChevronDown size={13} className={cn('transition-transform', isCollapsed && '-rotate-90')} />
                </button>
              )}
              {!isCollapsed && (
                <div className="space-y-0.5">
                  {items.map((item) => {
                    const Icon = iconMap[item.icon];
                    const active = isActive(item.matchPrefix);
                    return (
                      <Link
                        key={item.href}
                        href={item.href}
                        className={cn(
                          'flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-sm font-medium transition',
                          active
                            ? 'bg-white/12 text-cream'
                            : 'text-cream/55 hover:text-cream/90 hover:bg-white/8'
                        )}
                      >
                        {Icon && <Icon size={18} className={active ? 'opacity-100' : 'opacity-50'} />}
                        {item.label}
                      </Link>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </nav>
    </aside>
  );
}
