'use client';

import { use } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { ArrowLeft } from 'lucide-react';
import { SaleDetailPanel } from '@/components/sales/sale-detail-panel';

// 판매 상세 풀페이지 = 패널 컴포넌트 재사용(SSOT). 표시·발송·결제·수정 로직은 SaleDetailPanel 단일 출처.
export default function SaleDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();

  return (
    <>
      <Topbar title="판매 상세" />
      <div className="px-4 md:px-6 py-4 space-y-4 max-w-3xl">
        <Button variant="ghost" size="sm" onClick={() => router.push('/sales')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>
        <SaleDetailPanel saleId={id} />
      </div>
    </>
  );
}
