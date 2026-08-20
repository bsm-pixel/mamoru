'use client';

import { useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';
import { X } from 'lucide-react';
import { SerialPicker } from '@/components/sales/serial-picker';
import type { OrderItem } from '@/lib/supabase/types';
import type { OrderSerial } from '@/hooks/use-orders';

interface Props {
  orderId: string;
  items: OrderItem[];
  serials: OrderSerial[]; // 현재 배정된 시리얼
  onClose: () => void;
}

/**
 * 아임웹 주문 시리얼 배정 모달 — 판매의 SerialPicker 재사용(DRY).
 * 저장은 /api/orders/[id]/serials 로 "최종 배정 시리얼 id 집합"을 보내 서버가 add/remove·재고 치환 처리.
 */
export function OrderSerialModal({ orderId, items, serials, onClose }: Props) {
  const qc = useQueryClient();
  const [saving, setSaving] = useState(false);

  // product_id 있는 품목만 (시리얼 대상)
  const serialItems = items.filter((i) => i.product_id);

  // 품목별 선택 시리얼 id — 초기값 = 현재 배정된 시리얼
  const [sel, setSel] = useState<Record<string, string[]>>(() => {
    const m: Record<string, string[]> = {};
    for (const it of serialItems) {
      m[it.id] = serials.filter((s) => s.product_id === it.product_id).map((s) => s.id);
    }
    return m;
  });
  // 품목별 수동/자동생성 시리얼 번호 (재고에 없어 새로 만드는 것)
  const [manualSel, setManualSel] = useState<Record<string, string[]>>({});

  async function save() {
    if (saving) return;
    setSaving(true);
    try {
      const serialIds = Array.from(new Set(Object.values(sel).flat()));
      // 수동/자동생성 시리얼 번호 → product_id 별로 묶기
      const manualByProduct: Record<string, string[]> = {};
      for (const it of serialItems) {
        const arr = (manualSel[it.id] || []).map((s) => s.trim()).filter(Boolean);
        if (arr.length) {
          const pid = it.product_id as string;
          manualByProduct[pid] = [...(manualByProduct[pid] || []), ...arr];
        }
      }
      const res = await fetch(`/api/orders/${orderId}/serials`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ serialIds, manualByProduct }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error || '저장 실패');
      }
      const data = await res.json();
      toast.success(`시리얼 배정 완료 (추가 ${data.added} · 해제 ${data.removed})`);
      qc.invalidateQueries({ queryKey: ['order', orderId] });
      qc.invalidateQueries({ queryKey: ['orders'] });
      qc.invalidateQueries({ queryKey: ['products'] });
      qc.invalidateQueries({ queryKey: ['inventory'] });
      qc.invalidateQueries({ queryKey: ['serials'] });
      qc.invalidateQueries({ queryKey: ['serials-available'] });
      onClose();
    } catch (e) {
      toast.error(String(e instanceof Error ? e.message : e));
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-md max-h-[85vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-100">
          <h3 className="text-sm font-bold">시리얼 배정</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        <div className="flex-1 overflow-y-auto px-4 py-3 space-y-4">
          {serialItems.length === 0 ? (
            <p className="text-sm text-neutral-400 text-center py-6">시리얼 배정 가능한 품목이 없습니다</p>
          ) : (
            serialItems.map((it) => {
              const pid = it.product_id as string;
              const current = serials
                .filter((s) => s.product_id === pid)
                .map((s) => ({ id: s.id, serial_number: s.serial_number }));
              return (
                <div key={it.id}>
                  <div className="flex items-center justify-between mb-1">
                    <p className="text-sm font-medium truncate">{it.product_name}</p>
                    <span className="text-xs text-neutral-400 shrink-0 ml-2">{it.quantity}개</span>
                  </div>
                  <SerialPicker
                    productId={pid}
                    quantity={it.quantity}
                    selectedSerialIds={sel[it.id] || []}
                    onSelect={(ids) => setSel((prev) => ({ ...prev, [it.id]: ids }))}
                    currentSerials={current}
                    manualSerials={manualSel[it.id] || []}
                    onManualSerialsChange={(arr) => setManualSel((prev) => ({ ...prev, [it.id]: arr }))}
                  />
                </div>
              );
            })
          )}
        </div>

        <div className="px-4 py-3 border-t border-neutral-100 flex gap-2">
          <button onClick={onClose} className="flex-1 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600">취소</button>
          <button onClick={save} disabled={saving} className="flex-1 py-2 rounded-lg bg-neutral-900 text-white text-sm font-semibold disabled:opacity-50">
            {saving ? '저장 중...' : '저장'}
          </button>
        </div>
      </div>
    </div>
  );
}
