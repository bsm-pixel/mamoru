'use client';

import { Sidebar } from '@/components/layout/sidebar';
import { MobileNav } from '@/components/layout/mobile-nav';
import { usePushNotifications } from '@/hooks/use-push-notifications';

export default function DashboardLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  // 설정에서 알림이 켜져있으면 자동으로 FCM 구독
  usePushNotifications();

  return (
    <div className="flex min-h-screen bg-cream">
      <Sidebar />
      <main className="flex-1 flex flex-col min-w-0 pb-16 md:pb-0">
        {children}
      </main>
      <MobileNav />
    </div>
  );
}
