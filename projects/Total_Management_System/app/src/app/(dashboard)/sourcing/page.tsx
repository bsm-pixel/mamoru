'use client';

import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useSourcingList, useCreateSourcing } from '@/hooks/use-sourcing';
import { Plus, Sparkles, PackageSearch, CheckCircle2, XCircle } from 'lucide-react';

/**
 * 샘플 소싱 — 발주 목록.
 * 역할: 1688 샘플 소싱 → 입고매칭 → 실테스트 → 선별. (수량·SKU·products INSERT 없음)
 */
export default function SourcingListPage() {
  const router = useRouter();
  const { data, isLoading } = useSourcingList();
  const create = useCreateSourcing();

  const handleNew = async () => {
    const res = await create.mutateAsync({});
    router.push(`/sourcing/${res.id}`);
  };

  const orders = data?.orders ?? [];

  return (
    <div className="min-h-screen bg-warm-ivory">
      <Topbar title="샘플 소싱" />
      <div className="p-4 sm:p-6 max-w-5xl mx-auto space-y-4">
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <p className="text-sm text-neutral-600">
            1688 샘플을 들여와 <strong>입고매칭 → 실테스트 → 선별</strong>하는 도구. 선별된 제품은 리스트로 뽑아 직접 등록.
          </p>
          <Button onClick={handleNew} loading={create.isPending}>
            <Plus size={16} className="mr-1" /> 새 소싱
          </Button>
        </div>

        {isLoading ? (
          <div className="text-center py-16 text-sm text-neutral-400">불러오는 중…</div>
        ) : orders.length === 0 ? (
          <Card className="text-center py-16">
            <Sparkles className="mx-auto text-neutral-300 mb-3" size={32} />
            <p className="text-sm text-neutral-500 mb-4">아직 소싱 발주가 없습니다.</p>
            <Button onClick={handleNew} loading={create.isPending}>
              <Plus size={16} className="mr-1" /> 첫 소싱 시작
            </Button>
          </Card>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            {orders.map((po) => (
              <button
                key={po.id}
                type="button"
                onClick={() => router.push(`/sourcing/${po.id}`)}
                className="text-left rounded-xl border border-neutral-200 bg-white p-4 hover:border-neutral-400 hover:shadow-sm transition"
              >
                <div className="flex items-center justify-between mb-2">
                  <span className="font-mono text-sm font-bold text-indigo-black">{po.po_number}</span>
                  <span className="text-[11px] text-neutral-400">{po.order_date}</span>
                </div>
                <div className="text-sm text-neutral-700 truncate mb-3">
                  {po.memo || <span className="text-neutral-400">샘플 {po.counts.total}종</span>}
                </div>
                <div className="flex items-center gap-3 text-[11px]">
                  <span className="inline-flex items-center gap-1 text-neutral-500">
                    <PackageSearch size={12} /> 전체 {po.counts.total}
                  </span>
                  <span className="inline-flex items-center gap-1 text-emerald-600 font-medium">
                    <CheckCircle2 size={12} /> 채택 {po.counts.selected}
                  </span>
                  <span className="inline-flex items-center gap-1 text-rose-500">
                    <XCircle size={12} /> 탈락 {po.counts.rejected}
                  </span>
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
