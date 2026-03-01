import type { Metadata, Viewport } from 'next';

/* 간편진단 전용 PWA — manifest-diagnosis.json으로 오버라이드 */
export const metadata: Metadata = {
  title: 'MAMORU 간편진단',
  description: 'MAMORU 가위 간편진단',
  manifest: '/manifest-diagnosis.json',
  icons: { apple: '/icon-192.png' },
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  maximumScale: 1,
  themeColor: '#181725',
};

export default function DiagnosisLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return <>{children}</>;
}
