'use client';

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { EmptyState } from '@/components/ui/empty-state';
import { Package, Plus, ExternalLink, X } from 'lucide-react';
import toast from 'react-hot-toast';

interface Supply {
  id: string;
  name: string;
  sku: string;
  purchase_url: string | null;
  supply_status: string;
  image_url: string | null;
  price_purchase: number;
}

const STATUS_CONFIG: Record<string, { label: string; color: string; dot: string }> = {
  sufficient: { label: '충분', color: 'bg-green-100 text-green-700', dot: 'bg-green-500' },
  needed: { label: '주문필요', color: 'bg-red-100 text-red-700', dot: 'bg-red-500' },
  ordered: { label: '주문완료', color: 'bg-amber-100 text-amber-700', dot: 'bg-amber-500' },
};

const STATUS_CYCLE = ['sufficient', 'needed', 'ordered'];

export default function SuppliesPage() {
  const queryClient = useQueryClient();
  const [showAdd, setShowAdd] = useState(false);

  const { data, isLoading } = useQuery({
    queryKey: ['supplies'],
    queryFn: async () => {
      const res = await fetch('/api/supplies');
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ supplies: Supply[] }>;
    },
  });

  const updateStatus = useMutation({
    mutationFn: async ({ id, supply_status }: { id: string; supply_status: string }) => {
      const res = await fetch('/api/supplies', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id, supply_status }),
      });
      if (!res.ok) throw new Error(await res.text());
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ['supplies'] }),
  });

  const supplies = data?.supplies || [];
  const neededCount = supplies.filter((s) => s.supply_status === 'needed').length;
  const orderedCount = supplies.filter((s) => s.supply_status === 'ordered').length;

  function cycleStatus(supply: Supply) {
    const currentIdx = STATUS_CYCLE.indexOf(supply.supply_status);
    const nextStatus = STATUS_CYCLE[(currentIdx + 1) % STATUS_CYCLE.length];
    updateStatus.mutate({ id: supply.id, supply_status: nextStatus });
  }

  return (
    <>
      <Topbar title="부자재 관리" action={
        <Button size="sm" onClick={() => setShowAdd(true)}>
          <Plus size={14} />
          부자재 추가
        </Button>
      } />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 요약 */}
        {(neededCount > 0 || orderedCount > 0) && (
          <div className="flex gap-3">
            {neededCount > 0 && (
              <Badge className="bg-red-100 text-red-700">주문필요 {neededCount}건</Badge>
            )}
            {orderedCount > 0 && (
              <Badge className="bg-amber-100 text-amber-700">주문완료 대기 {orderedCount}건</Badge>
            )}
          </div>
        )}

        {/* 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 4 }).map((_, i) => <Skeleton key={i} className="h-16 w-full" />)}
            </div>
          ) : supplies.length === 0 ? (
            <EmptyState icon={Package} message="등록된 부자재가 없습니다" />
          ) : (
            <div className="divide-y divide-neutral-100">
              {supplies.map((s) => {
                const status = STATUS_CONFIG[s.supply_status] || STATUS_CONFIG.sufficient;
                return (
                  <div key={s.id} className="flex items-center gap-4 px-4 py-3 hover:bg-warm-ivory/60 transition">
                    {/* 상태 토글 */}
                    <button
                      onClick={() => cycleStatus(s)}
                      className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition flex items-center gap-1.5 ${status.color} hover:opacity-80`}
                      title="클릭하여 상태 변경"
                    >
                      <span className={`w-2 h-2 rounded-full ${status.dot}`} />
                      {status.label}
                    </button>

                    {/* 정보 */}
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-semibold text-indigo-black truncate">{s.name}</p>
                      <p className="text-xs text-neutral-400">{s.sku}</p>
                    </div>

                    {/* 주문 링크 */}
                    {s.purchase_url ? (
                      <a
                        href={s.purchase_url}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="shrink-0 flex items-center gap-1 px-3 py-1.5 rounded-lg bg-neutral-100 text-xs font-medium text-neutral-600 hover:bg-neutral-200 transition"
                      >
                        주문하기
                        <ExternalLink size={12} />
                      </a>
                    ) : (
                      <span className="text-xs text-neutral-300 shrink-0">링크 없음</span>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </Card>

        <p className="text-xs text-neutral-400">
          상태 뱃지를 클릭하면 충분 → 주문필요 → 주문완료 순서로 변경됩니다.
        </p>
      </div>

      {showAdd && <AddSupplyModal onClose={() => setShowAdd(false)} />}
    </>
  );
}

function AddSupplyModal({ onClose }: { onClose: () => void }) {
  const queryClient = useQueryClient();
  const [form, setForm] = useState({ name: '', purchase_url: '', memo: '' });
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit() {
    if (!form.name.trim()) { toast.error('부자재명을 입력해주세요'); return; }
    setSubmitting(true);
    try {
      const res = await fetch('/api/supplies', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      });
      if (!res.ok) throw new Error((await res.json()).error);
      toast.success('부자재 등록 완료');
      queryClient.invalidateQueries({ queryKey: ['supplies'] });
      onClose();
    } catch (err) {
      toast.error(String(err));
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl w-full max-w-md mx-4 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-100">
          <h2 className="text-sm font-bold text-indigo-black">부자재 추가</h2>
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center"><X size={16} /></button>
        </div>
        <div className="px-5 py-4 space-y-3">
          <div>
            <label className="text-xs text-neutral-500">부자재명 *</label>
            <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
              placeholder="예: 박스 (대), 뽁뽁이 50m" autoFocus
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
          </div>
          <div>
            <label className="text-xs text-neutral-500">주문 링크</label>
            <input type="url" value={form.purchase_url} onChange={(e) => setForm({ ...form, purchase_url: e.target.value })}
              placeholder="https://shopping.naver.com/..."
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
          </div>
          <div>
            <label className="text-xs text-neutral-500">메모</label>
            <input type="text" value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })}
              placeholder="규격, 수량 단위 등 (선택)"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
          </div>
        </div>
        <div className="px-5 py-4 border-t border-neutral-100 flex gap-2">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button className="flex-1" disabled={!form.name.trim() || submitting} onClick={handleSubmit}>
            {submitting ? '등록 중...' : '부자재 등록'}
          </Button>
        </div>
      </div>
    </div>
  );
}
