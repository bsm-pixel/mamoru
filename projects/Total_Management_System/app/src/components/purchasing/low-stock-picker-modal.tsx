'use client';

import { useState, useMemo } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { formatKRW } from '@/lib/utils/format';
import { AlertTriangle, Check } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';

interface LowStockPickerModalProps {
  open: boolean;
  onClose: () => void;
  products: Product[];
  threshold: number;
  alreadyAddedIds: Set<string>;
  onAdd: (products: Product[]) => void;
}

export function LowStockPickerModal({
  open, onClose, products, threshold, alreadyAddedIds, onAdd,
}: LowStockPickerModalProps) {
  const [checkedIds, setCheckedIds] = useState<Set<string>>(new Set());

  const lowStockProducts = useMemo(() =>
    products.filter((p) => p.stock_quantity >= 0 && p.stock_quantity <= threshold),
    [products, threshold]
  );

  const selectableProducts = useMemo(() =>
    lowStockProducts.filter((p) => !alreadyAddedIds.has(p.id)),
    [lowStockProducts, alreadyAddedIds]
  );

  const allSelected = selectableProducts.length > 0 && selectableProducts.every((p) => checkedIds.has(p.id));

  function toggleAll() {
    if (allSelected) {
      setCheckedIds(new Set());
    } else {
      setCheckedIds(new Set(selectableProducts.map((p) => p.id)));
    }
  }

  function toggleOne(id: string) {
    setCheckedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  function handleAdd() {
    const selected = lowStockProducts.filter((p) => checkedIds.has(p.id));
    onAdd(selected);
    setCheckedIds(new Set());
    onClose();
  }

  function handleClose() {
    setCheckedIds(new Set());
    onClose();
  }

  return (
    <Modal open={open} onClose={handleClose} title="저재고 상품 선택" className="max-w-xl">
      {/* 기준 안내 */}
      <p className="text-xs text-neutral-400 mb-3">
        재고 {threshold}개 이하 상품 · {lowStockProducts.length}건
      </p>

      {lowStockProducts.length === 0 ? (
        <p className="text-sm text-neutral-400 text-center py-8">저재고 상품이 없습니다</p>
      ) : (
        <>
          {/* 전체 선택 */}
          {selectableProducts.length > 0 && (
            <label className="flex items-center gap-2 px-3 py-2 rounded-lg bg-warm-ivory mb-2 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={allSelected}
                onChange={toggleAll}
                className="w-4 h-4 rounded accent-neutral-900"
              />
              <span className="text-xs font-semibold text-neutral-600">
                전체 선택 ({selectableProducts.length}개)
              </span>
            </label>
          )}

          {/* 상품 목록 */}
          <div className="max-h-[55vh] overflow-y-auto space-y-1">
            {lowStockProducts.map((p) => {
              const isAdded = alreadyAddedIds.has(p.id);
              const isChecked = isAdded || checkedIds.has(p.id);

              return (
                <label
                  key={p.id}
                  className={`flex items-center gap-3 px-3 py-2.5 rounded-lg border transition cursor-pointer select-none ${
                    isAdded
                      ? 'border-neutral-100 bg-neutral-50 opacity-50 cursor-default'
                      : isChecked
                        ? 'border-neutral-300 bg-neutral-50'
                        : 'border-transparent hover:bg-warm-ivory'
                  }`}
                >
                  <input
                    type="checkbox"
                    checked={isChecked}
                    disabled={isAdded}
                    onChange={() => !isAdded && toggleOne(p.id)}
                    className="w-4 h-4 rounded accent-neutral-900 flex-shrink-0"
                  />

                  {/* 제품 정보 */}
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-semibold text-indigo-black truncate">{p.name}</p>
                    <p className="text-xs text-neutral-500">{p.sku}</p>
                  </div>

                  {/* 재고 배지 */}
                  {isAdded ? (
                    <span className="flex items-center gap-1 text-xs text-neutral-400 flex-shrink-0">
                      <Check size={12} /> 추가됨
                    </span>
                  ) : (
                    <span className={`text-xs font-semibold px-2 py-0.5 rounded-full flex-shrink-0 ${
                      p.stock_quantity === 0
                        ? 'bg-red-50 text-red-600'
                        : 'bg-amber-50 text-amber-600'
                    }`}>
                      {p.stock_quantity}개
                    </span>
                  )}

                  {/* 매입가 */}
                  <span className="text-xs text-neutral-500 flex-shrink-0 w-16 text-right">
                    {p.price_purchase > 0 ? formatKRW(p.price_purchase) : '-'}
                  </span>
                </label>
              );
            })}
          </div>

          {/* 푸터 */}
          <div className="flex items-center justify-between mt-4 pt-3 border-t border-neutral-200">
            <span className="text-xs text-neutral-500">
              {checkedIds.size}개 선택됨
            </span>
            <div className="flex gap-2">
              <Button variant="ghost" size="sm" onClick={handleClose}>취소</Button>
              <Button size="sm" disabled={checkedIds.size === 0} onClick={handleAdd}>
                추가
              </Button>
            </div>
          </div>
        </>
      )}
    </Modal>
  );
}
