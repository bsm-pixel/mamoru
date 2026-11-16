'use client';

import { useEffect, useRef, useState } from 'react';
import { useRouter } from 'next/navigation';
import { createClient } from '@/lib/supabase/client';
import { Menu, LogOut, RefreshCw, User, Search, ClipboardList } from 'lucide-react';
import { openUnifiedPrep } from '@/components/shared/unified-prep-host';
import { openQuickFind } from '@/components/layout/quick-find';
import { KeyHint } from '@/components/ui/key-hint';
import toast from 'react-hot-toast';
import { usePageTitle } from '@/hooks/use-page-title';

interface TopbarProps {
  title: string;
  /** 이 화면에서 처리할 건수 — 브라우저 탭 이름에 「화면 이름 (3) · TMS」로 붙는다 */
  count?: number | null;
  onMenuToggle?: () => void;
  action?: React.ReactNode;
  /** 페이지가 자체 새로고침 버튼을 action 으로 줄 때 공통 새로고침 아이콘을 숨긴다 (아이콘 2개 중복 방지) */
  hideRefresh?: boolean;
}

export function Topbar({ title, count, onMenuToggle, action, hideRefresh = false }: TopbarProps) {
  const router = useRouter();
  usePageTitle(title, count);

  // 2026-11-15 로그아웃이 새로고침 바로 옆이라 잘못 누르기 쉬웠다 → 사람 아이콘 안으로
  const [menuOpen, setMenuOpen] = useState(false);
  const [email, setEmail] = useState<string | null>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!menuOpen) return;
    if (email === null) {
      createClient().auth.getUser().then(({ data }) => setEmail(data.user?.email || '')).catch(() => setEmail(''));
    }
    const onDown = (e: MouseEvent) => { if (!menuRef.current?.contains(e.target as Node)) setMenuOpen(false); };
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') setMenuOpen(false); };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => { document.removeEventListener('mousedown', onDown); document.removeEventListener('keydown', onKey); };
  }, [menuOpen, email]);

  const handleLogout = async () => {
    const supabase = createClient();
    await supabase.auth.signOut();
    toast.success('로그아웃 완료');
    router.push('/login');
    router.refresh();
  };

  return (
    <header className="sticky top-0 z-30 flex items-center justify-between h-14 px-4 md:px-6 bg-card-white/90 backdrop-blur-sm border-b border-neutral-200">
      <div className="flex items-center gap-3">
        {/* 모바일 메뉴 버튼 */}
        <button title="메뉴" aria-label="메뉴"
          onClick={onMenuToggle}
          className="md:hidden w-9 h-9 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
        >
          <Menu size={20} className="text-neutral-700" />
        </button>
        <h2 className="text-base font-bold text-indigo-black">{title}</h2>
      </div>

      <div className="flex items-center gap-2">
        {action}
        {/* 2026-11-15 빠른 찾기 — 고객 · 판매번호 · 제품 · 메뉴 */}
        <button
          type="button"
          onClick={openQuickFind}
          className="h-9 px-2.5 lg:px-3 flex items-center gap-1.5 rounded-lg border border-neutral-200 text-xs text-neutral-500 hover:bg-warm-ivory transition"
          title="빠른 찾기 (Ctrl+K)" aria-label="빠른 찾기"
        >
          <Search size={14} /><span className="hidden lg:inline">찾기</span><KeyHint>Ctrl K</KeyHint>
        </button>
        {/* 2026-11-16 전체 출고 준비표 — 전엔 주문 · 판매 · 복원수리 세 화면에 따로 있던 버튼을 여기 하나로 */}
        <button
          type="button"
          onClick={openUnifiedPrep}
          className="h-9 px-2.5 lg:px-3 flex items-center gap-1.5 rounded-lg border border-neutral-200 text-xs text-neutral-500 hover:bg-warm-ivory transition"
          title="전체 출고 준비표 — 복원수리·주문·판매의 출고 준비를 한 장에" aria-label="전체 출고 준비표"
        >
          <ClipboardList size={14} /><span className="hidden lg:inline">출고 준비표</span>
        </button>
        {!hideRefresh && <button
          onClick={() => router.refresh()}
          className="w-9 h-9 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
          title="새로고침"
        >
          <RefreshCw size={16} className="text-neutral-500" />
        </button>}
        <div ref={menuRef} className="relative">
          <button
            type="button"
            onClick={() => setMenuOpen((o) => !o)}
            aria-haspopup="menu" aria-expanded={menuOpen}
            className="w-9 h-9 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
            title="내 계정" aria-label="내 계정"
          >
            <User size={16} className="text-neutral-500" />
          </button>
          {menuOpen && (
            <div role="menu" className="absolute right-0 top-full mt-1 w-52 rounded-xl border border-neutral-200 bg-white p-1.5 z-40">
              {email && <p className="px-2.5 py-1.5 text-[11px] text-neutral-400 truncate">{email}</p>}
              <button
                type="button" role="menuitem"
                onClick={handleLogout}
                className="w-full h-9 px-2.5 flex items-center gap-2 rounded-lg text-sm font-semibold text-neutral-700 hover:bg-warm-ivory transition"
              >
                <LogOut size={14} className="text-neutral-500" />로그아웃
              </button>
            </div>
          )}
        </div>
      </div>
    </header>
  );
}
