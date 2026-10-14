'use client';

import { useState } from 'react';
import { useQuery, useQueryClient } from '@tanstack/react-query';
import { Card } from '@/components/ui/card';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import toast from 'react-hot-toast';

interface ReturnStockData {
  serials: Array<{ id: string; serial_number: string; updated_at: string }>;
  return_stock: number;
}

/** 처리 대상 — 시리얼 1개 또는 비시리얼 1개 */
interface Target {
  serialId?: string;
  label: string;
  action: 'restock' | 'defective';
}

/**
 * 반품창고 카드 (2026-10-14) — 제품 상세 패널.
 * 반품·교환으로 회수된 제품을 실물 확인 후 [재판매] 또는 [불량]으로 꺼낸다.
 * 재판매 = 현재고·아임웹 +1, 시리얼은 판매대기/진열 창고로 → 판매 입력 시리얼 목록에 뜸.
 * 반품창고가 비어 있으면 카드 자체를 숨긴다.
 */
export function ReturnStockCard({ productId }: { productId: string }) {
  const queryClient = useQueryClient();
  const [target, setTarget] = useState<Target | null>(null);
  const [zone, setZone] = useState<'ready' | 'display'>('ready');

  const { data } = useQuery<ReturnStockData>({
    queryKey: ['return-stock', productId],
    queryFn: async () => {
      const res = await fetch(`/api/inventory/return-release?product_id=${productId}`);
      if (!res.ok) throw new Error('반품창고 조회 실패');
      return res.json();
    },
    enabled: !!productId,
  });

  const serials = data?.serials || [];
  const nonSerial = data?.return_stock || 0;
  if (serials.length === 0 && nonSerial === 0) return null;

  async function run() {
    if (!target) return;
    const res = await fetch('/api/inventory/return-release', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ product_id: productId, serial_id: target.serialId, action: target.action, zone }),
    });
    const json = await res.json();
    if (!res.ok) {
      toast.error(json.error || '처리 실패');
      return;
    }
    toast.success(target.action === 'restock' ? `${target.label} → 재판매 가능 (현재고 +1)` : `${target.label} → 불량 처리`);
    for (const key of [['return-stock', productId], ['product', productId], ['products'], ['inventory'], ['serials'], ['serials-available']]) {
      queryClient.invalidateQueries({ queryKey: key });
    }
    setTarget(null);
  }

  const rowBtn = 'text-xs px-2.5 py-1 rounded-md border transition';

  return (
    <Card>
      <h4 className="text-xs text-neutral-500 mb-1">반품창고 · 검수 대기 {serials.length + nonSerial}개</h4>
      <p className="text-[11px] text-neutral-400 mb-3">실물 확인 후 처리하세요. 판매가능 현재고·아임웹에는 아직 반영되지 않은 상태입니다.</p>
      <div className="space-y-2">
        {serials.map((s) => (
          <div key={s.id} className="flex items-center justify-between gap-2">
            <span className="text-sm font-mono text-neutral-800 truncate">{s.serial_number}</span>
            <div className="flex gap-1.5 shrink-0">
              <button className={`${rowBtn} border-neutral-900 bg-neutral-900 text-white hover:bg-neutral-700`}
                onClick={() => { setZone('ready'); setTarget({ serialId: s.id, label: s.serial_number, action: 'restock' }); }}>
                재판매
              </button>
              <button className={`${rowBtn} border-neutral-200 text-neutral-500 hover:bg-neutral-50`}
                onClick={() => setTarget({ serialId: s.id, label: s.serial_number, action: 'defective' })}>
                불량
              </button>
            </div>
          </div>
        ))}
        {nonSerial > 0 && (
          <div className="flex items-center justify-between gap-2">
            <span className="text-sm text-neutral-800">시리얼 없음 {nonSerial}개</span>
            <div className="flex gap-1.5 shrink-0">
              <button className={`${rowBtn} border-neutral-900 bg-neutral-900 text-white hover:bg-neutral-700`}
                onClick={() => setTarget({ label: '반품 1개', action: 'restock' })}>
                1개 재판매
              </button>
              <button className={`${rowBtn} border-neutral-200 text-neutral-500 hover:bg-neutral-50`}
                onClick={() => setTarget({ label: '반품 1개', action: 'defective' })}>
                1개 불량
              </button>
            </div>
          </div>
        )}
      </div>

      <ConfirmModal
        open={!!target}
        onClose={() => setTarget(null)}
        onConfirm={run}
        variant={target?.action === 'defective' ? 'danger' : 'default'}
        title={target?.action === 'restock' ? '재판매 가능으로 전환' : '불량 처리'}
        confirmLabel={target?.action === 'restock' ? '재판매 전환' : '불량 처리'}
        message={target?.action === 'restock' ? (
          <div className="space-y-3 text-sm">
            <p><b>{target.label}</b>을(를) 판매가능 재고로 돌립니다. 현재고와 아임웹 재고가 1개 늘어납니다.</p>
            {target.serialId && (
              <div className="flex gap-2">
                {(['ready', 'display'] as const).map((z) => (
                  <button key={z} type="button" onClick={() => setZone(z)}
                    className={`flex-1 text-xs py-2 rounded-md border ${zone === z ? 'border-neutral-900 bg-neutral-900 text-white' : 'border-neutral-200 text-neutral-600'}`}>
                    {z === 'ready' ? '판매대기 창고로' : '진열 창고로'}
                  </button>
                ))}
              </div>
            )}
          </div>
        ) : (
          <p className="text-sm"><b>{target?.label}</b>을(를) 불량으로 처리합니다. 판매가능 재고에는 포함되지 않습니다.</p>
        )}
      />
    </Card>
  );
}
