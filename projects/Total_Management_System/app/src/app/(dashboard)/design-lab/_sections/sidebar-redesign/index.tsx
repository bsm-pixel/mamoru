'use client';

import { useState } from 'react';
import { cn } from '@/lib/utils/cn';
import { NAV_GROUPS } from '@/lib/utils/constants';
import {
  LayoutDashboard, ShoppingCart, MessageSquare, CalendarOff, Wrench, Store, Truck,
  FileSignature, Users, Handshake, Package, Boxes, PackageSearch, PackageOpen,
  Search, Star, BarChart3, Wallet, Building2, Settings, ChevronDown, Circle,
  type LucideIcon,
} from 'lucide-react';

const iconMap: Record<string, LucideIcon> = {
  LayoutDashboard, ShoppingCart, MessageSquare, CalendarOff, Wrench, Store, Truck,
  FileSignature, Users, Handshake, Package, Boxes, PackageSearch, PackageOpen,
  Search, Star, BarChart3, Wallet, Building2, Settings,
};

const DEMO_ACTIVE = '/sourcing'; // 데모 active 하이라이트

type Variant = 'dark' | 'light';

/**
 * § 사이드바 리디자인 시안 — 디자인모니터 비교용 (실제 네비 무변경).
 * 다크/라이트 2종 × 그룹 접기(아코디언). 결정 후 실제 sidebar.tsx 반영 + § 삭제.
 */
export function SidebarRedesignSection() {
  return (
    <section className="space-y-5">
      <div className="rounded-2xl border-2 border-amber-300 bg-amber-50 p-5">
        <div className="flex items-center gap-2 mb-2">
          <span className="text-base font-bold text-stone-900">§ 사이드바 리디자인</span>
          <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-200 text-amber-900 uppercase tracking-wider">시안 비교</span>
        </div>
        <p className="text-xs text-stone-700 leading-relaxed">
          <strong>다크 vs 라이트</strong> 두 톤 + <strong>그룹 접기(아코디언)</strong>. 그룹 헤더(▾) 클릭하면 접힘/펼침.
          마음에 드는 쪽 고르시면 실제 사이드바에 반영합니다.
          <br />
          <span className="text-stone-500">※ 사이드바 <strong>화면 고정(스크롤해도 안 사라짐)</strong>은 이미 실제 적용됨. 이 시안은 "디자인"만 비교.</span>
        </p>
      </div>

      <div className="flex flex-col lg:flex-row gap-6 items-start">
        <Frame label="🌑 다크 (정체성 유지)"><MiniSidebar variant="dark" /></Frame>
        <Frame label="☀️ 라이트 (노션/리니어풍)"><MiniSidebar variant="light" /></Frame>

        <div className="flex-1 min-w-0 text-xs text-stone-600 space-y-2 leading-relaxed">
          <div className="font-bold text-stone-900">공통 개선점</div>
          <ul className="list-disc ml-4 space-y-1">
            <li><strong>그룹 접기</strong>: 22개 메뉴를 평소 접어두고 필요할 때만 펼침 → 깔끔</li>
            <li><strong>둥근 pill 액티브</strong>: 현재 좌측 막대선 → 부드러운 pill 하이라이트</li>
            <li>여백·타이포 정돈, 아이콘 일관성</li>
            <li>화면 <strong>고정</strong>(스크롤해도 항상 보임) — 이미 적용</li>
          </ul>
          <div className="font-bold text-stone-900 pt-1">톤 차이</div>
          <ul className="list-disc ml-4 space-y-1">
            <li><strong>다크</strong>: 마모루 "조용히 압도" 톤 유지, 변화 작음·안전</li>
            <li><strong>라이트</strong>: 최신 SaaS 느낌·가벼움, 본문(크림)과 매끄럽게 이어짐</li>
          </ul>
        </div>
      </div>
    </section>
  );
}

function Frame({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="flex-shrink-0">
      <div className="text-[11px] font-bold text-stone-500 mb-2">{label}</div>
      <div className="rounded-xl border border-stone-300 shadow-sm overflow-hidden" style={{ width: 232, height: 600 }}>
        {children}
      </div>
    </div>
  );
}

function MiniSidebar({ variant }: { variant: Variant }) {
  const dark = variant === 'dark';
  // 비어있지 않은 그룹 인덱스만 아코디언 대상. 기본 전부 펼침.
  const [collapsed, setCollapsed] = useState<Record<number, boolean>>({});
  const toggle = (gi: number) => setCollapsed((c) => ({ ...c, [gi]: !c[gi] }));

  return (
    <div className={cn('h-full flex flex-col overflow-y-auto', dark ? 'bg-indigo-black text-cream' : 'bg-card-white text-indigo-black')}>
      {/* 로고 */}
      <div className="px-4 py-4 flex-shrink-0">
        <h1 className="text-base font-extrabold tracking-tight">MAMORU</h1>
        <p className={cn('text-[10px] mt-0.5', dark ? 'text-cream/45' : 'text-neutral-400')}>TMS v1.0</p>
      </div>

      <nav className="flex-1 px-2.5 pb-4 space-y-1">
        {NAV_GROUPS.map((group, gi) => {
          const hasHeader = !!group.group;
          const isCollapsed = hasHeader && collapsed[gi];
          return (
            <div key={gi}>
              {hasHeader && (
                <button
                  type="button"
                  onClick={() => toggle(gi)}
                  className={cn(
                    'w-full flex items-center justify-between px-2.5 mt-2 mb-1 group',
                    dark ? 'text-cream/45 hover:text-cream/70' : 'text-neutral-400 hover:text-neutral-600'
                  )}
                >
                  <span className="text-[10px] font-semibold tracking-wider uppercase">{group.group}</span>
                  <ChevronDown size={12} className={cn('transition-transform', isCollapsed && '-rotate-90')} />
                </button>
              )}
              {!isCollapsed && (
                <div className="space-y-0.5">
                  {group.items.map((item) => {
                    const Icon = iconMap[item.icon] || Circle;
                    const active = item.matchPrefix === DEMO_ACTIVE;
                    return (
                      <div
                        key={item.href}
                        className={cn(
                          'flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-[13px] font-medium transition cursor-default',
                          active
                            ? dark
                              ? 'bg-white/12 text-cream'
                              : 'bg-neutral-100 text-indigo-black font-semibold'
                            : dark
                              ? 'text-cream/60 hover:bg-white/8'
                              : 'text-neutral-500 hover:bg-neutral-100'
                        )}
                      >
                        <Icon size={17} className={cn(active ? 'opacity-100' : dark ? 'opacity-50' : 'text-neutral-400')} />
                        {item.label}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </nav>
    </div>
  );
}
