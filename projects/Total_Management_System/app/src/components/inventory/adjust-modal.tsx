'use client';

import { useState } from 'react';
import { Button } from '@/components/ui/button';
import { X, Plus, Minus } from 'lucide-react';
import type { InventoryItem } from '@/hooks/use-inventory';
import toast from 'react-hot-toast';

const ADJUSTMENT_TYPES = [
  { value: 'correction', label: '실사 보정' },
  { value: 'damage', label: '파손/불량' },
  { value: 'return', label: '반품 입고' },
  { value: 'other', label: '기타' },
];

interface AdjustModalProps {
  items: InventoryItem[];
  onClose: () => void;
  onSuccess: () => void;
}

export function AdjustModal({ items, onClose, onSuccess }: AdjustModalProps) {
  const [productId, setProductId] = useState('');
  const [adjustmentType, setAdjustmentType] = useState('correction');
  const [quantity, setQuantity] = useState(0);
  const [reason, setReason] = useState('');
  const [submitting, setSubmitting] = useState(false);

  const selectedProduct = items.find((i) => i.id === productId);
  const newQty = selectedProduct ? selectedProduct.stock_quantity + quantity : 0;
  const isValid = productId && quantity !== 0 && newQty >= 0;

  async function handleSubmit() {
    if (!isValid) return;
    setSubmitting(true);
    try {
      const res = await fetch('/api/inventory/adjust', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ product_id: productId, adjustment_type: adjustmentType, quantity, reason }),
      });
      if (!res.ok) {
        const data = await res.json();
        throw new Error(data.error || '조정 실패');
      }
      const data = await res.json();
      toast.success(`${data.product_name} 재고 ${data.previous_qty} → ${data.new_qty}개`);
      onSuccess();
      onClose();
    } catch (err) {
      toast.error(String(err));
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div
        className="bg-white rounded-xl w-full max-w-md mx-4 shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-100">
          <h2 className="text-sm font-bold text-indigo-black">재고 조정</h2>
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center">
            <X size={16} />
          </button>
        </div>

        {/* 본문 */}
        <div className="px-5 py-4 space-y-4">
          {/* 제품 선택 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">제품 *</label>
            <select
              value={productId}
              onChange={(e) => { setProductId(e.target.value); setQuantity(0); }}
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            >
              <option value="">제품 선택...</option>
              {items.map((i) => (
                <option key={i.id} value={i.id}>
                  {i.name} ({i.sku}) — 현재 {i.stock_quantity}개
                </option>
              ))}
            </select>
          </div>

          {/* 조정 유형 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">조정 유형</label>
            <div className="flex gap-1.5 flex-wrap">
              {ADJUSTMENT_TYPES.map((t) => (
                <button
                  key={t.value}
                  onClick={() => setAdjustmentType(t.value)}
                  className={`px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                    adjustmentType === t.value
                      ? 'bg-terracotta text-cream'
                      : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>
          </div>

          {/* 수량 조정 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">수량 변경 *</label>
            <div className="flex items-center gap-3">
              <button
                onClick={() => setQuantity((q) => q - 1)}
                disabled={!selectedProduct || newQty <= 0}
                className="w-10 h-10 rounded-lg bg-red-50 text-red-500 flex items-center justify-center hover:bg-red-100 disabled:opacity-30"
              >
                <Minus size={18} />
              </button>
              <input
                type="number"
                value={quantity}
                onChange={(e) => setQuantity(parseInt(e.target.value) || 0)}
                className="w-20 h-10 text-center rounded-lg border border-neutral-200 bg-warm-ivory text-lg font-bold focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
              <button
                onClick={() => setQuantity((q) => q + 1)}
                disabled={!selectedProduct}
                className="w-10 h-10 rounded-lg bg-green-50 text-green-600 flex items-center justify-center hover:bg-green-100 disabled:opacity-30"
              >
                <Plus size={18} />
              </button>
            </div>
            {selectedProduct && (
              <p className="text-xs text-neutral-400 mt-1.5">
                현재 <strong>{selectedProduct.stock_quantity}</strong>개 →{' '}
                <strong className={newQty < 0 ? 'text-red-500' : quantity > 0 ? 'text-green-600' : quantity < 0 ? 'text-red-500' : ''}>
                  {newQty}
                </strong>개
                {quantity !== 0 && (
                  <span className={quantity > 0 ? 'text-green-600' : 'text-red-500'}>
                    {' '}({quantity > 0 ? '+' : ''}{quantity})
                  </span>
                )}
              </p>
            )}
            {newQty < 0 && (
              <p className="text-xs text-red-500 mt-1">재고가 음수가 될 수 없습니다</p>
            )}
          </div>

          {/* 사유 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">사유 (선택)</label>
            <input
              type="text"
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              placeholder="조정 사유를 입력하세요"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            />
          </div>
        </div>

        {/* 푸터 */}
        <div className="px-5 py-4 border-t border-neutral-100 flex gap-2">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button
            className="flex-1"
            disabled={!isValid || submitting}
            onClick={handleSubmit}
          >
            {submitting ? '처리 중...' : '재고 조정'}
          </Button>
        </div>
      </div>
    </div>
  );
}
