'use client';

import { useRepair } from '@/hooks/use-repairs';
import { CustomerDetailPanel } from '@/components/customers/customer-detail-panel';
import { SkeletonRows } from '@/components/ui/skeleton';

/**
 * 큰 화면(가로 2200 이상)의 세 번째 칸 — 고른 복원수리 건의 고객 이력 · 메모 (2026-11-15)
 * 넓은 화면을 표를 늘리는 데 쓰지 않고 "이 고객이 전에 무엇을 했는지"를 같이 보여 준다.
 * useRepair 는 상세 칸과 같은 자료(같은 queryKey)라 따로 더 받지 않는다.
 */
export function RepairCustomerAside({ repairId }: { repairId: string }) {
  const { data, isLoading } = useRepair(repairId);
  if (isLoading) return <SkeletonRows rows={5} />;
  const customerId = data?.repair?.customer_id;
  if (!customerId) {
    return <p className="p-6 text-xs text-neutral-400 text-center">이 건은 고객 정보와 연결돼 있지 않습니다</p>;
  }
  return <CustomerDetailPanel customerId={customerId} />;
}
