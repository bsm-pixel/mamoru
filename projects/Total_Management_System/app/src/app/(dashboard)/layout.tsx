'use client';

import Head from 'next/head';
import { Sidebar } from '@/components/layout/sidebar';
import { MobileNav } from '@/components/layout/mobile-nav';

export default function DashboardLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <>
      <head>
        <link rel="manifest" href="/manifest.json" />
      </head>
      <div className="flex min-h-screen bg-cream">
        <Sidebar />
        <main className="flex-1 flex flex-col min-w-0 pb-16 md:pb-0">
          {children}
        </main>
        <MobileNav />
      </div>
    </>
  );
}
