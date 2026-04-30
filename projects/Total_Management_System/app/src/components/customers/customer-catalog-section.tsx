'use client';

import { useState } from 'react';
import { Plus, Search, Save, Trash2 } from 'lucide-react';
import { formatKRW } from '@/lib/utils/format';
import { useProducts } from '@/hooks/use-sales';
import { useCustomerCatalog, useAddToCustomerCatalog, useUpdateCustomerCatalog, useRemoveFromCustomerCatalog } from '@/hooks/use-customer-catalog';

/**
 * B2B 납품처(dealer/academy) 납품품목 카탈로그 UI
 * SupplierCatalogSection과 mirror 패턴 (마이그 073)
 *
 * 동작:
 *  - 제품에서 불러오기 → catalog 추가
 *  - 각 항목에 delivery_name(납품명) + features(특징) 등록
 *  - delivery_name 비어있으면 product.name fallback (sales 페이지 addProduct에서)
 */
export function CustomerCatalogSection({ customerId }: { customerId: string }) {
  const { data, isLoading } = useCustomerCatalog(customerId);
  const { data: products = [] } = useProducts();
  const addToCatalog = useAddToCustomerCatalog();
  const updateCatalog = useUpdateCustomerCatalog();
  const removeFromCatalog = useRemoveFromCustomerCatalog();
  const [showProductPicker, setShowProductPicker] = useState(false);
  const [productSearch, setProductSearch] = useState('');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState({ delivery_name: '', features: '', unit_price: '' });

  const catalog = data?.catalog || [];
  const catalogProductIds = new Set(catalog.map((c) => c.product_id));
  const availableProducts = products.filter((p) => !catalogProductIds.has(p.id) && p.category !== 'SUP');
  const filtered = productSearch.length >= 1
    ? availableProducts.filter((p) => p.name.toLowerCase().includes(productSearch.toLowerCase()) || (p.sku || '').toLowerCase().includes(productSearch.toLowerCase())).slice(0, 10)
    : availableProducts.slice(0, 10);

  function startEdit(entry: typeof catalog[0]) {
    setEditingId(entry.id);
    setEditForm({
      delivery_name: entry.delivery_name,
      features: entry.features,
      unit_price: entry.unit_price != null ? String(entry.unit_price) : '',
    });
  }

  async function saveEdit() {
    if (!editingId) return;
    const parsedPrice = editForm.unit_price.trim() ? parseInt(editForm.unit_price.replace(/[^0-9]/g, ''), 10) : null;
    const unitPrice = parsedPrice != null && !isNaN(parsedPrice) && parsedPrice > 0 ? parsedPrice : null;
    await updateCatalog.mutateAsync({
      customerId,
      catalogId: editingId,
      deliveryName: editForm.delivery_name,
      features: editForm.features,
      unitPrice,
    });
    setEditingId(null);
  }

  async function handleAdd(productId: string) {
    await addToCatalog.mutateAsync({ customerId, productIds: [productId] });
    setShowProductPicker(false);
    setProductSearch('');
  }

  async function handleRemove(catalogId: string) {
    await removeFromCatalog.mutateAsync({ customerId, catalogId });
  }

  return (
    <div className="px-4 py-4 space-y-3">
      {/* 상단 버튼 */}
      <div className="flex items-center justify-between">
        <p className="text-xs font-semibold text-neutral-600">납품품목 ({catalog.length})</p>
        <button
          onClick={() => setShowProductPicker(!showProductPicker)}
          className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-neutral-900 text-white text-xs font-medium hover:bg-neutral-800 transition"
        >
          <Plus size={12} />제품에서 불러오기
        </button>
      </div>

      {/* 제품 선택 드롭다운 */}
      {showProductPicker && (
        <div className="rounded-lg border border-neutral-200 bg-white p-3 space-y-2">
          <div className="relative">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              type="text"
              value={productSearch}
              onChange={(e) => setProductSearch(e.target.value)}
              placeholder="제품명 또는 SKU 검색"
              className="w-full h-8 pl-8 pr-3 rounded-lg border border-neutral-200 text-xs"
              autoFocus
            />
          </div>
          <div className="max-h-40 overflow-y-auto space-y-0.5">
            {filtered.length === 0 ? (
              <p className="text-xs text-neutral-400 text-center py-3">추가 가능한 제품 없음</p>
            ) : filtered.map((p) => (
              <button
                key={p.id}
                onClick={() => handleAdd(p.id)}
                className="w-full flex items-center justify-between px-3 py-2 rounded hover:bg-neutral-50 text-left"
              >
                <div>
                  <span className="text-xs font-medium">{p.name}</span>
                  {p.sku && !p.sku.startsWith('IW-') && <span className="text-[10px] text-neutral-400 ml-2">{p.sku}</span>}
                </div>
                {p.price > 0 && <span className="text-[10px] text-neutral-500">{formatKRW(p.price)}</span>}
              </button>
            ))}
          </div>
          <button onClick={() => { setShowProductPicker(false); setProductSearch(''); }} className="w-full py-1.5 text-xs text-neutral-400 hover:text-neutral-600">닫기</button>
        </div>
      )}

      {/* 카탈로그 목록 */}
      {isLoading ? (
        <div className="text-xs text-neutral-400 text-center py-8">로딩중...</div>
      ) : catalog.length === 0 ? (
        <div className="text-xs text-neutral-400 text-center py-8">등록된 납품품목이 없습니다<br />위 버튼으로 제품을 추가해주세요</div>
      ) : (
        <div className="space-y-2">
          {catalog.map((entry) => (
            <div key={entry.id} className="rounded-lg border border-neutral-200 p-3">
              {editingId === entry.id ? (
                /* 편집 모드 */
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold">{entry.product_name}</span>
                    <span className="text-[10px] text-neutral-400">정가 {formatKRW(entry.price)}</span>
                  </div>
                  <div className="grid grid-cols-3 gap-2">
                    <div>
                      <label className="text-[10px] text-neutral-400">납품명</label>
                      <input value={editForm.delivery_name} onChange={(e) => setEditForm({ ...editForm, delivery_name: e.target.value })}
                        placeholder="송장/납품서 출력용 품명"
                        className="w-full h-7 px-2 rounded border border-neutral-200 text-xs" />
                    </div>
                    <div>
                      <label className="text-[10px] text-neutral-400">납품가</label>
                      <input value={editForm.unit_price} onChange={(e) => setEditForm({ ...editForm, unit_price: e.target.value })}
                        placeholder="비우면 정가/그룹가"
                        inputMode="numeric"
                        className="w-full h-7 px-2 rounded border border-neutral-200 text-xs" />
                    </div>
                    <div>
                      <label className="text-[10px] text-neutral-400">특징</label>
                      <input value={editForm.features} onChange={(e) => setEditForm({ ...editForm, features: e.target.value })}
                        placeholder="규격, 특이사항 등"
                        className="w-full h-7 px-2 rounded border border-neutral-200 text-xs" />
                    </div>
                  </div>
                  <div className="flex gap-1.5 justify-end">
                    <button onClick={() => setEditingId(null)} className="px-2 py-1 text-[10px] text-neutral-500 hover:text-neutral-700">취소</button>
                    <button onClick={saveEdit} className="flex items-center gap-1 px-2 py-1 rounded bg-neutral-900 text-white text-[10px]">
                      <Save size={10} />저장
                    </button>
                  </div>
                </div>
              ) : (
                /* 읽기 모드 — 납품명을 헤더로 (송장에 출력될 이름) */
                <div className="flex items-start justify-between">
                  <div className="flex-1 min-w-0" onClick={() => startEdit(entry)} style={{ cursor: 'pointer' }}>
                    <div className="flex items-baseline gap-2 mb-1 flex-wrap">
                      {entry.delivery_name ? (
                        <>
                          <span className="text-sm font-bold text-blue-700 truncate">{entry.delivery_name}</span>
                          <span className="text-[10px] text-neutral-400">송장 출력명</span>
                        </>
                      ) : (
                        <span className="text-sm font-semibold text-neutral-400 italic">납품명 미입력 → 클릭하여 등록</span>
                      )}
                    </div>
                    <div className="flex gap-3 text-[11px] text-neutral-500 flex-wrap">
                      <span className="text-neutral-400">TMS 제품: <span className="text-neutral-600">{entry.product_name}</span></span>
                      <span>정가 {formatKRW(entry.price)}</span>
                      <span>납품가: {entry.unit_price != null ? <span className="text-blue-600 font-semibold">{formatKRW(entry.unit_price)}</span> : <em className="text-neutral-300">정가 사용</em>}</span>
                      {entry.features && <span>특징: {entry.features}</span>}
                    </div>
                  </div>
                  <button onClick={() => handleRemove(entry.id)} className="text-neutral-300 hover:text-red-500 ml-2 shrink-0">
                    <Trash2 size={12} />
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
