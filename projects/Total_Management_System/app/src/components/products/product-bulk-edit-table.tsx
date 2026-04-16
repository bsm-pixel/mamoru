'use client';

import { useState, useCallback, useRef } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';
import { ArrowLeft, Save } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { usePriceGroups } from '@/hooks/use-price-groups';
import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CAT_LABELS } from '@/lib/utils/setting-defaults';
import type { Product } from '@/lib/supabase/types';

/* ── 변경 추적 타입 ── */
type EditedFields = {
  name?: string;
  category?: string;
  price?: number;
  price_purchase?: number;
  purchase_name?: string;
  // price_groups 하위: { [groupKey]: { price?, display_name? } }
  price_groups?: Record<string, { price?: number; display_name?: string }>;
};

interface Props {
  products: Product[];
  onClose: () => void;
}

/* ── 직접 PATCH (훅 미사용 — 중간 invalidate 방지) ── */
async function patchProduct(id: string, data: Record<string, unknown>) {
  const res = await fetch(`/api/products/${id}`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err));
  }
  return res.json();
}

export function ProductBulkEditTable({ products, onClose }: Props) {
  const queryClient = useQueryClient();
  const priceGroups = usePriceGroups();
  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const categories = useSetting<string[]>('inventory.categories', Object.keys(DEFAULT_CAT_LABELS));

  const [editedRows, setEditedRows] = useState<Map<string, EditedFields>>(new Map());
  const [saving, setSaving] = useState(false);
  const [saveProgress, setSaveProgress] = useState({ done: 0, total: 0 });
  const [failedIds, setFailedIds] = useState<Set<string>>(new Set());
  const tableRef = useRef<HTMLDivElement>(null);

  const groupKeys = Object.keys(priceGroups); // e.g. ['dealer', 'academy']
  const dirtyCount = editedRows.size;

  /* ── 셀 변경 핸들러 ── */
  const updateField = useCallback((productId: string, field: keyof EditedFields, value: unknown) => {
    setEditedRows(prev => {
      const next = new Map(prev);
      const existing = next.get(productId) || {};
      const product = products.find(p => p.id === productId);
      if (!product) return prev;

      const updated = { ...existing, [field]: value };

      // 원본과 같으면 해당 필드 삭제
      const original = (product as Record<string, unknown>)[field];
      if (value === original || (typeof value === 'number' && value === Number(original))) {
        delete (updated as Record<string, unknown>)[field];
      }

      // 빈 객체면 행 자체 삭제
      if (Object.keys(updated).length === 0) {
        next.delete(productId);
      } else {
        next.set(productId, updated);
      }
      return next;
    });
  }, [products]);

  /* ── 가격 그룹 셀 변경 ── */
  const updateGroupField = useCallback((productId: string, groupKey: string, subField: 'price' | 'display_name', value: unknown) => {
    setEditedRows(prev => {
      const next = new Map(prev);
      const existing = next.get(productId) || {};
      const product = products.find(p => p.id === productId);
      if (!product) return prev;

      const existingGroups = existing.price_groups || {};
      const existingGroup = existingGroups[groupKey] || {};
      const updatedGroup = { ...existingGroup, [subField]: value };

      // 원본과 비교
      const origGroup = product.price_groups?.[groupKey];
      const origValue = subField === 'price'
        ? (origGroup?.price ?? 0)
        : (origGroup?.display_name ?? '');
      const newValue = subField === 'price' ? Number(value) || 0 : value;
      if (newValue === origValue) {
        delete (updatedGroup as Record<string, unknown>)[subField];
      }

      // 빈 그룹이면 삭제
      const updatedGroups = { ...existingGroups };
      if (Object.keys(updatedGroup).length === 0) {
        delete updatedGroups[groupKey];
      } else {
        updatedGroups[groupKey] = updatedGroup;
      }

      const updated = { ...existing };
      if (Object.keys(updatedGroups).length === 0) {
        delete updated.price_groups;
      } else {
        updated.price_groups = updatedGroups;
      }

      if (Object.keys(updated).length === 0) {
        next.delete(productId);
      } else {
        next.set(productId, updated);
      }
      return next;
    });
  }, [products]);

  /* ── 셀 값 읽기 (편집값 우선 → 원본 fallback) ── */
  const getCellValue = (product: Product, field: keyof EditedFields) => {
    const edited = editedRows.get(product.id);
    if (edited && field in edited) return (edited as Record<string, unknown>)[field];
    return (product as Record<string, unknown>)[field];
  };

  const getGroupValue = (product: Product, groupKey: string, subField: 'price' | 'display_name') => {
    const edited = editedRows.get(product.id);
    if (edited?.price_groups?.[groupKey]?.[subField] !== undefined) {
      return edited.price_groups[groupKey][subField];
    }
    const orig = product.price_groups?.[groupKey];
    if (subField === 'price') return orig?.price ?? 0;
    return orig?.display_name ?? '';
  };

  /* ── 일괄 저장 ── */
  const handleSave = async () => {
    const entries = Array.from(editedRows.entries());
    if (entries.length === 0) { toast('변경사항이 없습니다'); return; }

    setSaving(true);
    setSaveProgress({ done: 0, total: entries.length });
    setFailedIds(new Set());

    let successCount = 0;
    const newFailed = new Set<string>();

    for (const [productId, changes] of entries) {
      try {
        const product = products.find(p => p.id === productId);
        const payload: Record<string, unknown> = {};

        if (changes.name !== undefined) payload.name = changes.name;
        if (changes.category !== undefined) payload.category = changes.category;
        if (changes.price !== undefined) payload.price = changes.price;
        if (changes.price_purchase !== undefined) payload.price_purchase = changes.price_purchase;
        if (changes.purchase_name !== undefined) payload.purchase_name = changes.purchase_name;

        // price_groups: 원본 머지 + dual-write
        if (changes.price_groups) {
          const merged: Record<string, { price?: number; display_name?: string }> = {};
          // 원본 복사
          if (product?.price_groups) {
            for (const [k, v] of Object.entries(product.price_groups)) {
              merged[k] = { price: v.price ?? undefined, display_name: v.display_name ?? undefined };
            }
          }
          // 변경분 덮어쓰기
          for (const [k, v] of Object.entries(changes.price_groups)) {
            merged[k] = { ...merged[k], ...v };
          }
          payload.price_groups = merged;
          // dual-write
          payload.price_dealer = merged['dealer']?.price || 0;
          payload.price_academy = merged['academy']?.price || 0;
          payload.dealer_name = merged['dealer']?.display_name || null;
          payload.academy_name = merged['academy']?.display_name || null;
        }

        await patchProduct(productId, payload);
        successCount++;
      } catch (err) {
        newFailed.add(productId);
        console.error(`[bulk-edit] ${productId} 실패:`, err);
      }
      setSaveProgress(prev => ({ ...prev, done: prev.done + 1 }));
    }

    setSaving(false);
    setFailedIds(newFailed);

    if (newFailed.size === 0) {
      toast.success(`${successCount}개 제품 수정 완료`);
      setEditedRows(new Map());
    } else {
      toast.error(`${successCount}개 성공, ${newFailed.size}개 실패`);
      // 성공한 항목은 editedRows에서 제거
      setEditedRows(prev => {
        const next = new Map(prev);
        for (const id of prev.keys()) {
          if (!newFailed.has(id)) next.delete(id);
        }
        return next;
      });
    }

    queryClient.invalidateQueries({ queryKey: ['products'] });
  };

  /* ── 닫기 (미저장 확인) ── */
  const handleClose = () => {
    if (dirtyCount > 0 && !confirm(`${dirtyCount}개 변경사항이 저장되지 않았습니다. 나가시겠습니까?`)) return;
    onClose();
  };

  /* ── 공통 input 스타일 ── */
  const inputCls = 'w-full h-8 px-2 text-sm border-0 bg-transparent focus:bg-blue-50 focus:ring-1 focus:ring-blue-300 focus:outline-none rounded';
  const numCls = `${inputCls} text-right font-mono tabular-nums`;
  const readonlyCls = 'w-full h-8 px-2 text-xs font-mono text-neutral-400 truncate flex items-center';

  return (
    <div className="flex flex-col h-full min-h-0">
      {/* ── 상단 툴바 ── */}
      <div className="shrink-0 flex items-center gap-3 px-4 py-3 border-b border-neutral-200 bg-white">
        <button onClick={handleClose} className="flex items-center gap-1 text-sm text-neutral-500 hover:text-neutral-800">
          <ArrowLeft size={16} />
          목록으로
        </button>
        <div className="flex-1 text-center">
          <span className="text-sm font-semibold text-neutral-700">일괄 수정</span>
          <span className="ml-2 text-xs text-neutral-400">({products.length}개 제품)</span>
        </div>
        <div className="flex items-center gap-2">
          {dirtyCount > 0 && (
            <span className="px-2 py-0.5 text-xs font-medium rounded-full bg-amber-100 text-amber-700">
              변경 {dirtyCount}건
            </span>
          )}
          {saving && (
            <span className="text-xs text-neutral-500">
              {saveProgress.done}/{saveProgress.total} 저장 중...
            </span>
          )}
          <Button size="sm" onClick={handleSave} disabled={saving || dirtyCount === 0}>
            <Save size={14} />
            {saving ? '저장 중...' : '저장'}
          </Button>
        </div>
      </div>

      {/* ── 테이블 ── */}
      <div ref={tableRef} className="flex-1 min-h-0 overflow-auto">
        <table className="w-max min-w-full border-collapse text-sm">
          <thead className="sticky top-0 z-10 bg-neutral-50 border-b border-neutral-200">
            <tr>
              <th className="sticky left-0 z-20 bg-neutral-50 w-2" /> {/* 변경 인디케이터 */}
              <th className="sticky left-2 z-20 bg-neutral-50 px-2 py-2 text-left text-xs font-medium text-neutral-500 w-[80px]">SKU</th>
              <th className="sticky left-[98px] z-20 bg-neutral-50 px-2 py-2 text-left text-xs font-medium text-neutral-500 w-[180px] border-r border-neutral-200">제품명</th>
              <th className="px-2 py-2 text-left text-xs font-medium text-neutral-500 w-[100px]">카테고리</th>
              <th className="px-2 py-2 text-right text-xs font-medium text-neutral-500 w-[100px]">소매가</th>
              {groupKeys.map(gk => (
                <th key={`price-${gk}`} className="px-2 py-2 text-right text-xs font-medium text-neutral-500 w-[100px]">
                  {priceGroups[gk]?.label || gk}
                </th>
              ))}
              <th className="px-2 py-2 text-right text-xs font-medium text-neutral-500 w-[100px]">매입가</th>
              <th className="px-2 py-2 text-left text-xs font-medium text-neutral-500 w-[150px]">발주명</th>
              {groupKeys.map(gk => (
                <th key={`dn-${gk}`} className="px-2 py-2 text-left text-xs font-medium text-neutral-500 w-[140px]">
                  {priceGroups[gk]?.label || gk} 납품명
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {products.map(p => {
              const isDirty = editedRows.has(p.id);
              const isFailed = failedIds.has(p.id);
              return (
                <tr
                  key={p.id}
                  className={`border-b border-neutral-100 ${isDirty ? 'bg-amber-50/50' : 'hover:bg-neutral-50/50'} ${isFailed ? 'ring-1 ring-red-300' : ''} ${!p.is_active ? 'opacity-50' : ''}`}
                >
                  {/* 변경 인디케이터 */}
                  <td className="sticky left-0 z-10 w-2 p-0 bg-inherit">
                    {isDirty && <div className="w-1 h-full bg-amber-400 rounded-r" />}
                  </td>
                  {/* SKU (읽기전용) */}
                  <td className="sticky left-2 z-10 bg-inherit">
                    <div className={readonlyCls}>{p.sku}</div>
                  </td>
                  {/* 제품명 */}
                  <td className="sticky left-[98px] z-10 bg-inherit border-r border-neutral-200">
                    <input
                      className={inputCls}
                      value={getCellValue(p, 'name') as string || ''}
                      onChange={e => updateField(p.id, 'name', e.target.value)}
                    />
                  </td>
                  {/* 카테고리 */}
                  <td>
                    <select
                      className={`${inputCls} cursor-pointer`}
                      value={getCellValue(p, 'category') as string || ''}
                      onChange={e => updateField(p.id, 'category', e.target.value)}
                    >
                      {categories.map(cat => (
                        <option key={cat} value={cat}>{catLabels[cat] || cat}</option>
                      ))}
                    </select>
                  </td>
                  {/* 소매가 */}
                  <td>
                    <input
                      type="number"
                      className={numCls}
                      value={getCellValue(p, 'price') as number ?? 0}
                      onChange={e => updateField(p.id, 'price', Number(e.target.value) || 0)}
                    />
                  </td>
                  {/* 그룹별 가격 */}
                  {groupKeys.map(gk => (
                    <td key={`price-${gk}`}>
                      <input
                        type="number"
                        className={numCls}
                        value={getGroupValue(p, gk, 'price') as number ?? 0}
                        onChange={e => updateGroupField(p.id, gk, 'price', Number(e.target.value) || 0)}
                      />
                    </td>
                  ))}
                  {/* 매입가 */}
                  <td>
                    <input
                      type="number"
                      className={numCls}
                      value={getCellValue(p, 'price_purchase') as number ?? 0}
                      onChange={e => updateField(p.id, 'price_purchase', Number(e.target.value) || 0)}
                    />
                  </td>
                  {/* 발주명 */}
                  <td>
                    <input
                      className={inputCls}
                      value={getCellValue(p, 'purchase_name') as string || ''}
                      onChange={e => updateField(p.id, 'purchase_name', e.target.value)}
                    />
                  </td>
                  {/* 그룹별 납품명 */}
                  {groupKeys.map(gk => (
                    <td key={`dn-${gk}`}>
                      <input
                        className={inputCls}
                        value={getGroupValue(p, gk, 'display_name') as string || ''}
                        onChange={e => updateGroupField(p.id, gk, 'display_name', e.target.value)}
                      />
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
