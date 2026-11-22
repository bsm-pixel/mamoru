'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { useProducts } from '@/hooks/use-sales';
import { formatKRW } from '@/lib/utils/format';
import type { Product } from '@/lib/supabase/types';

const CAT_LABEL: Record<string, string> = {
  BL: '블런트',
  TH: '틴닝',
  LO: '장가위',
  SL: '슬라이싱',
};

interface Props {
  open: boolean;
  onClose: () => void;
  onSelect: (product: Product) => void;
}

export function ProductPickerModal({ open, onClose, onSelect }: Props) {
  const { data: products = [] } = useProducts();
  const [category, setCategory] = useState('all');

  const categories = Array.from(new Set(products.map((p) => p.category))).sort();
  const filtered = category === 'all' ? products : products.filter((p) => p.category === category);

  function handleSelect(product: Product) {
    onSelect(product);
    onClose();
  }

  return (
    <Modal open={open} onClose={onClose} title="제품 선택">
      {/* 카테고리 탭 */}
      <div className="flex gap-2 mb-4 overflow-x-auto pb-1">
        <button
          onClick={() => setCategory('all')}
          className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
            category === 'all' ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500'
          }`}
        >
          전체
        </button>
        {categories.map((cat) => (
          <button
            key={cat}
            onClick={() => setCategory(cat)}
            className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
              category === cat ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500'
            }`}
          >
            {CAT_LABEL[cat] || cat}
          </button>
        ))}
      </div>

      {/* 제품 리스트 */}
      <div className="max-h-[50vh] overflow-y-auto -mx-5 px-5">
        {filtered.length === 0 ? (
          <p className="text-center text-sm text-neutral-400 py-8">제품이 없습니다</p>
        ) : (
          <div className="space-y-1">
            {filtered.map((p) => {
              // 2026-11-22 새 제품 재고(판매 가능 수량) — 없는 모델을 계약하지 않게. 고를 수는 있다(입고 예정 등)
              const stock = p.stock_quantity ?? 0;
              return (
                <button
                  key={p.id}
                  onClick={() => handleSelect(p)}
                  data-stock={stock}
                  className={`w-full text-left px-4 py-4 rounded-lg hover:bg-neutral-50 active:bg-neutral-100 transition border border-transparent hover:border-neutral-200 flex items-center gap-3 ${stock <= 0 ? 'opacity-55' : ''}`}
                >
                  <span className="min-w-0 flex-1">
                    <span className="block text-sm font-semibold text-neutral-800">{p.name}</span>
                    <span className="block text-xs text-neutral-500 mt-0.5">{p.sku} · {formatKRW(p.price)}</span>
                  </span>
                  <span className={`shrink-0 text-xs font-semibold tabular-nums whitespace-nowrap ${stock <= 0 ? 'text-red-600' : 'text-neutral-700'}`}>
                    {stock <= 0 ? '재고 없음' : `새 제품 ${stock}개`}
                  </span>
                </button>
              );
            })}
          </div>
        )}
      </div>
    </Modal>
  );
}
