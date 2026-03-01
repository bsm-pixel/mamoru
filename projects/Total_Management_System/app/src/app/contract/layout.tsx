import type { Metadata, Viewport } from 'next';

/* 계약서 전용 PWA — manifest-contract.json 사용 */
export const metadata: Metadata = {
  title: 'MAMORU 계약서',
  description: 'MAMORU 구매 계약서 작성',
  icons: { apple: '/icon-192.png' },
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  maximumScale: 1,
  themeColor: '#181725',
};

export default function ContractLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <>
      <head>
        <link rel="manifest" href="/manifest-contract.json" />
      </head>
      {children}
    </>
  );
}
