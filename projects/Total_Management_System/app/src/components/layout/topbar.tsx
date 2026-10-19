'use client';

import { useRouter } from 'next/navigation';
import { createClient } from '@/lib/supabase/client';
import { Menu, LogOut, RefreshCw } from 'lucide-react';
import toast from 'react-hot-toast';

interface TopbarProps {
  title: string;
  onMenuToggle?: () => void;
  action?: React.ReactNode;
  /** 페이지가 자체 새로고침 버튼을 action 으로 줄 때 공통 새로고침 아이콘을 숨긴다 (아이콘 2개 중복 방지) */
  hideRefresh?: boolean;
}

export function Topbar({ title, onMenuToggle, action, hideRefresh = false }: TopbarProps) {
  const router = useRouter();

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
        <button
          onClick={onMenuToggle}
          className="md:hidden w-9 h-9 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
        >
          <Menu size={20} className="text-neutral-700" />
        </button>
        <h2 className="text-base font-bold text-indigo-black">{title}</h2>
      </div>

      <div className="flex items-center gap-2">
        {action}
        {!hideRefresh && <button
          onClick={() => router.refresh()}
          className="w-9 h-9 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
          title="새로고침"
        >
          <RefreshCw size={16} className="text-neutral-500" />
        </button>}
        <button
          onClick={handleLogout}
          className="w-9 h-9 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
          title="로그아웃"
        >
          <LogOut size={16} className="text-neutral-500" />
        </button>
      </div>
    </header>
  );
}
