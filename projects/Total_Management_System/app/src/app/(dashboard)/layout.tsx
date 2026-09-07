'use client';

import { useRouter } from 'next/navigation';
import { Sidebar } from '@/components/layout/sidebar';
import { MobileNav } from '@/components/layout/mobile-nav';
import { ShortcutHelp } from '@/components/ui/shortcut-help';
import { usePushNotifications } from '@/hooks/use-push-notifications';
import { useHotkeys } from '@/hooks/use-hotkeys';

export default function DashboardLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  // 설정에서 알림이 켜져있으면 자동으로 FCM 구독
  usePushNotifications();

  // 전역 화면 이동 단축키 — Ctrl+[ 뒤로 / Ctrl+] 앞으로.
  // allowInInput 기본 false → 입력창에 타이핑 중엔 동작 안 함(폼 작성 중 실수 이탈 방지 = "화면에만")
  const router = useRouter();
  useHotkeys([
    { combo: 'ctrl+[', handler: () => router.back() },
    { combo: 'ctrl+]', handler: () => router.forward() },
  ]);

  return (
    <div className="flex min-h-screen bg-cream">
      <Sidebar />
      <main className="flex-1 flex flex-col min-w-0 pb-[calc(3.5rem+env(safe-area-inset-bottom))] md:pb-0">
        {children}
      </main>
      <MobileNav />
      <ShortcutHelp />
    </div>
  );
}
