'use client';

import { MamoruWordmark } from '@/components/ui/mamoru-logo';
import { useState, useEffect } from 'react';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import { LayoutDashboard, Star, ChevronDown, Settings } from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { NAV_GROUPS } from '@/lib/utils/constants';
import { NAV_ICON_MAP } from '@/lib/utils/nav-icons';
import { useSetting } from '@/hooks/use-settings';

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
  // 2026-11-15 MY MENU 에 있는 메뉴는 아래 전체 메뉴에서 빼고(중복 제거), 나머지는 「더 보기」로 접는다.
  //   지금 보고 있는 화면이 MY MENU 에 없으면 자동으로 펴서 어디에 있는지 보이게 한다.
  const favSet = new Set(favorites.map((f) => f.href));
  const toggleGroup = (key: string) => {
    setCollapsed((c) => {
      const next = { ...c, [key]: !c[key] };
      try { localStorage.setItem(COLLAPSE_KEY, JSON.stringify(next)); } catch { /* ignore */ }
      return next;
    });
  };

  // 「더 보기」 펼침 여부는 그룹 접힘과 같은 저장 칸에 함께 둔다(따로 읽어 오는 코드를 늘리지 않으려고)
  const moreOpen = !!collapsed.__more;
  const toggleMore = () => toggleGroup('__more');

  // matchPrefix 정확 매칭 (하위 경로가 별도 메뉴인 경우 겹치지 않게 — IA 정리 2026-05-18)
  const isActive = (matchPrefix: string) =>
    matchPrefix === '/consultations'
      ? pathname === '/consultations' || (pathname.startsWith('/consultations/') && !pathname.startsWith('/consultations/calendar'))
      : pathname.startsWith(matchPrefix);

  const dashActive = pathname.startsWith('/dashboard');
  const onFavorite = favorites.some((f) => isActive(f.matchPrefix));
  const settingsActive = pathname.startsWith('/settings');
  // 즐겨찾기가 없으면 접을 이유가 없다 → 항상 펼침
  const showAll = favorites.length === 0 || moreOpen || (!onFavorite && !dashActive && !settingsActive);

  return (
    <aside className="hidden md:flex flex-col w-56 h-screen sticky top-0 self-start bg-indigo-black text-cream border-r border-white/8">
      {/* 로고 + 대시보드(홈) 아이콘 — 대시보드는 메뉴 목록에서 빼고 여기로 */}
      <div className="px-4 py-4 flex items-center justify-between flex-shrink-0">
        <div>
          <h1 className="leading-none"><MamoruWordmark height={13} /></h1>
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

      {/* MY MENU — 즐겨찾기 (스크롤 밖 상단 고정, 내부 스크롤 없음 · 평면 리스트) */}
      {favorites.length > 0 && (
        <div className="flex-shrink-0 px-2.5 pt-1">
          <div className="rounded-xl bg-cream p-1.5 shadow-sm">
            <div className="flex items-center gap-1.5 px-1.5 pt-0.5 pb-1.5 text-indigo-black/50">
              <Star size={12} className="fill-current text-amber-500" />
              <span className="text-[11px] font-bold tracking-wider uppercase">MY MENU</span>
            </div>
            <div className="space-y-0.5">
              {favorites.map((item) => {
                const Icon = NAV_ICON_MAP[item.icon];
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
        </div>
      )}

      {/* 그룹별 내비 (그룹 접기 + pill 액티브) — 나머지, 스크롤 */}
      <nav className="flex-1 px-2.5 pb-2 overflow-y-auto">
        {favorites.length > 0 && (
          <button
            type="button"
            onClick={toggleMore}
            aria-expanded={showAll}
            className="w-full flex items-center justify-between px-2.5 mt-3 py-1.5 rounded-lg text-cream/60 hover:text-cream/90 hover:bg-white/8 transition"
          >
            <span className="text-xs font-semibold">{showAll ? '메뉴 접기' : '더 보기'}</span>
            <ChevronDown size={14} className={cn('transition-transform', showAll && 'rotate-180')} />
          </button>
        )}
        {showAll && NAV_GROUPS.map((group, gi) => {
          // 대시보드 · 숨김 · MY MENU 에 이미 있는 것 · 설정(맨 아래 고정) 제외
          const items = group.items.filter((it) => it.href !== '/dashboard' && it.href !== '/settings' && !hidden.has(it.href) && !favSet.has(it.href));
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
                    const Icon = NAV_ICON_MAP[item.icon];
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
      {/* 설정은 맨 아래 고정 (MY MENU 에 넣었으면 위에만) */}
      {!favSet.has('/settings') && !hidden.has('/settings') && (
        <div className="flex-shrink-0 px-2.5 pb-3 pt-1 border-t border-white/8">
          <Link
            href="/settings"
            className={cn(
              'flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-sm font-medium transition',
              settingsActive ? 'bg-white/12 text-cream' : 'text-cream/55 hover:text-cream/90 hover:bg-white/8'
            )}
          >
            <Settings size={18} className={settingsActive ? 'opacity-100' : 'opacity-50'} />
            설정
          </Link>
        </div>
      )}
    </aside>
  );
}
