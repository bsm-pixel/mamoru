'use client';

import { useRouter } from 'next/navigation';
import { Sidebar } from '@/components/layout/sidebar';
import { MobileNav } from '@/components/layout/mobile-nav';
import { ShortcutHelp } from '@/components/ui/shortcut-help';
import { usePushNotifications } from '@/hooks/use-push-notifications';
import { useHotkeys, isDialogOpen } from '@/hooks/use-hotkeys';
import { askBeforeLeave } from '@/hooks/use-unsaved-guard';
import { ListKeys } from '@/components/layout/list-keys';
import { QuickFind } from '@/components/layout/quick-find';
import { InputKeys } from '@/components/layout/input-keys';
import { OfflineBanner } from '@/components/layout/offline-banner';

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
  // 2026-11-15 F2(판매 입력) · F4(거래처 매출)를 어느 화면에서나. 기능키라 입력 칸에서도 동작.
  //   창이 떠 있으면 무시하고, 저장하지 않은 내용이 있으면 먼저 묻는다(판매 입력 · 발주 작성 · 계약서).
  const goByKey = (to: string) => {
    if (isDialogOpen()) return;
    if (window.location.pathname + window.location.search === to) return;
    if (askBeforeLeave(to)) return;
    router.push(to);
  };
  useHotkeys([
    { combo: 'ctrl+[', handler: () => router.back() },
    { combo: 'ctrl+]', handler: () => router.forward() },
    { combo: 'f2', allowInInput: true, handler: () => goByKey('/sales/new') },
    { combo: 'f4', allowInInput: true, handler: () => goByKey('/sales/new?mode=b2b') },
  ]);

  return (
    <div className="flex min-h-screen bg-cream">
      <Sidebar />
      <main className="flex-1 flex flex-col min-w-0 pb-[calc(3.5rem+env(safe-area-inset-bottom))] md:pb-0">
        {children}
      </main>
      <MobileNav />
      <ShortcutHelp />
      <ListKeys />
      <QuickFind />
      <InputKeys />
      <OfflineBanner />
    </div>
  );
}
