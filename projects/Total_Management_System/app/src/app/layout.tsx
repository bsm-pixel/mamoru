import type { Metadata, Viewport } from 'next';
import './globals.css';
import { Toaster } from 'react-hot-toast';
import { QueryProvider } from '@/components/providers/query-provider';
import { NotificationNavigator } from '@/components/notification-navigator';

export const metadata: Metadata = {
  title: 'MAMORU TMS',
  description: 'MAMORU 통합관리시스템',
  manifest: '/manifest.json',
  icons: { apple: '/icon-192.png' },
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  maximumScale: 1,
  themeColor: '#181725',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ko">
      <head>
        <link
          rel="stylesheet"
          as="style"
          crossOrigin="anonymous"
          href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css"
        />
      </head>
      <body className="antialiased">
        <QueryProvider>
          {children}
          <NotificationNavigator />
        </QueryProvider>
        <Toaster
          position="top-right"
          toastOptions={{
            style: {
              background: 'var(--mm-card-white)',
              color: 'var(--mm-indigo-black)',
              border: '1px solid var(--mm-neutral-200)',
            },
          }}
        />
      </body>
    </html>
  );
}
