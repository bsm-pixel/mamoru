'use client';

import { use } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { ArrowLeft } from 'lucide-react';
import { CustomerDetailPanel } from '@/components/customers/customer-detail-panel';

// 고객 상세 풀페이지 = 패널 컴포넌트 재사용(SSOT). 표시·편집 로직은 CustomerDetailPanel 단일 출처.
export default function CustomerDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();

  return (
    <>
      <Topbar title="고객 상세" />
      <div className="px-4 md:px-6 py-4 space-y-4 max-w-3xl">
        <Button variant="ghost" size="sm" onClick={() => router.push('/customers')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>
        <CustomerDetailPanel customerId={id} hideDetailLink />
      </div>
    </>
  );
}
