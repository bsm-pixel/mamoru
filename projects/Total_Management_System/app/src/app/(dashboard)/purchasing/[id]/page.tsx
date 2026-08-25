'use client';

import { use } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { ArrowLeft } from 'lucide-react';
import { PurchaseDetailPanel } from '@/components/purchasing/purchase-detail-panel';

// 발주 상세 풀페이지 = 패널 컴포넌트 재사용(SSOT). 편집·분할입고·검수·인쇄 로직은 PurchaseDetailPanel 단일 출처.
export default function PurchaseOrderDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();

  return (
    <>
      <Topbar title="발주 상세" />
      {/* 패널 최상위가 flex-1 가정 → flex 컨테이너로 감싸 정상 렌더 */}
      <div className="px-4 md:px-6 py-4 flex flex-col gap-4 max-w-3xl min-h-[calc(100vh-3.5rem)]">
        <Button variant="ghost" size="sm" className="self-start" onClick={() => router.push('/purchasing')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>
        <PurchaseDetailPanel purchaseId={id} />
      </div>
    </>
  );
}
