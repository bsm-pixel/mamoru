'use client';

import { useState } from 'react';
import Link from 'next/link';
import { usePathname } from 'next/navigation';
import {
  LayoutDashboard, ShoppingCart, MessageSquare, Wrench, Store,
  FileSignature, Package, Settings, Users, Truck, BarChart3,
  MoreHorizontal, X, Boxes, Star,
  type LucideIcon,
} from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { NAV_ITEMS } from '@/lib/utils/constants';

const iconMap: Record<string, LucideIcon> = {
  LayoutDashboard, ShoppingCart, MessageSquare, Wrench, Store,
  FileSignature, Package, Settings, Users, Truck, BarChart3, Boxes, Star,
};

// 모바일 하단 탭에 고정 표시할 4개 + 더보기
const MOBILE_TAB_PREFIXES = ['/dashboard', '/sales', '/consultations', '/repairs'];

export function MobileNav() {
  const pathname = usePathname();
  const [showMore, setShowMore] = useState(false);

  const mobileTabs = NAV_ITEMS.filter((item) =>
    MOBILE_TAB_PREFIXES.includes(item.matchPrefix)
  );
  const moreItems = NAV_ITEMS.filter((item) =>
    !MOBILE_TAB_PREFIXES.includes(item.matchPrefix)
  );

  // "더보기" 메뉴 안의 항목이 활성화되어있는지 확인
  const moreActive = moreItems.some((item) => pathname.startsWith(item.matchPrefix));

  return (
    <>
      {/* 더보기 바텀시트 */}
      {showMore && (
        <div className="md:hidden fixed inset-0 z-50">
          <div
            className="absolute inset-0 bg-black/30"
            onClick={() => setShowMore(false)}
          />
          <div className="absolute bottom-0 left-0 right-0 bg-card-white rounded-t-2xl safe-area-bottom animate-in slide-in-from-bottom duration-200">
            <div className="flex items-center justify-between px-4 pt-4 pb-2">
              <span className="text-sm font-bold text-indigo-black">메뉴</span>
              <button
                onClick={() => setShowMore(false)}
                className="w-7 h-7 rounded-full bg-neutral-100 flex items-center justify-center"
              >
                <X size={14} />
              </button>
            </div>
            <div className="grid grid-cols-4 gap-1 px-3 pb-4">
              {moreItems.map((item) => {
                const Icon = iconMap[item.icon];
                const active = pathname.startsWith(item.matchPrefix);
                return (
                  <Link
                    key={item.href}
                    href={item.href}
                    onClick={() => setShowMore(false)}
                    className={cn(
                      'flex flex-col items-center gap-1 py-3 rounded-xl transition',
                      active ? 'bg-terracotta/10 text-terracotta' : 'text-neutral-500 hover:bg-neutral-50'
                    )}
                  >
                    {Icon && <Icon size={20} />}
                    <span className="text-[10px] font-medium">{item.label}</span>
                  </Link>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* 하단 탭 바 (5개) */}
      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-card-white/95 backdrop-blur-sm border-t border-neutral-200 safe-area-bottom">
        <div className="flex items-center justify-around h-14">
          {mobileTabs.map((item) => {
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
          {/* 더보기 */}
          <button
            onClick={() => setShowMore(true)}
            className={cn(
              'flex flex-col items-center gap-0.5 min-w-[64px] py-1 transition',
              moreActive ? 'text-terracotta' : 'text-neutral-400'
            )}
          >
            <MoreHorizontal size={20} />
            <span className="text-[10px] font-medium">더보기</span>
          </button>
        </div>
      </nav>
    </>
  );
}
