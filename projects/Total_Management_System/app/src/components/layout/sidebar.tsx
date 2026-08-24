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
  CalendarDays,
  Zap,
  Tag,
  ChevronDown,
  type LucideIcon,
} from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { NAV_GROUPS } from '@/lib/utils/constants';
import { useSetting } from '@/hooks/use-settings';

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
  CalendarDays,
  Zap,
  Tag,
};

const COLLAPSE_KEY = 'tms-sidebar-collapsed';

export function Sidebar() {
  const pathname = usePathname();
  // 사이드바 커스텀 설정 (즐겨찾기 MY MENU + 숨김) — 설정>화면 설정에서 관리
  const cfg = useSetting<{ hidden?: string[]; favorites?: string[] }>('system.sidebar_config', {});
  const hidden = new Set(cfg.hidden || []);
  const itemByHref = Object.fromEntries(NAV_GROUPS.flatMap((g) => g.items).map((i) => [i.href, i]));
  const favorites = (cfg.favorites || [])
    .map((href) => itemByHref[href])
    .filter((it): it is (typeof NAV_GROUPS)[number]['items'][number] => !!it && it.href !== '/dashboard');
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
        {/* MY MENU — 즐겨찾기 바로가기 (설정>화면 설정에서 등록·순서). 은은한 앰버 톤 패널로 구분 */}
        {favorites.length > 0 && (
          <div className="mb-2 mt-1 rounded-xl bg-cream p-1.5 shadow-sm">
            <div className="flex items-center gap-1.5 px-1.5 pt-0.5 pb-1.5 text-indigo-black/50">
              <Star size={12} className="fill-current text-amber-500" />
              <span className="text-[11px] font-bold tracking-wider uppercase">MY MENU</span>
            </div>
            <div className="space-y-0.5">
              {favorites.map((item) => {
                const Icon = iconMap[item.icon];
                const active = isActive(item.matchPrefix);
                return (
                  <Link
                    key={`fav-${item.href}`}
                    href={item.href}
                    className={cn(
                      'flex items-center gap-2.5 px-2 py-2 rounded-lg text-sm font-medium transition',
                      active ? 'bg-indigo-black text-cream' : 'text-indigo-black/70 hover:bg-black/5 hover:text-indigo-black'
                    )}
                  >
                    {Icon && <Icon size={18} className={active ? 'opacity-100' : 'opacity-55'} />}
                    {item.label}
                  </Link>
                );
              })}
            </div>
          </div>
        )}

        {NAV_GROUPS.map((group, gi) => {
          const items = group.items.filter((it) => it.href !== '/dashboard' && !hidden.has(it.href)); // 대시보드·숨김 제외
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
