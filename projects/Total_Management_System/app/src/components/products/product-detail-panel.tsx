'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { useQueryClient } from '@tanstack/react-query';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { Modal } from '@/components/ui/modal';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { SupplierSelect } from '@/components/ui/supplier-select';
import { useProduct, useUpdateProduct, useCreateProduct, useSourcingByProduct } from '@/hooks/use-product-detail';
import { formatKRW } from '@/lib/utils/format';
import { Save, Package, Hash, X, Plus, Archive, Copy, Eye, EyeOff, Trash2, ArrowRightLeft, ArrowRight, ExternalLink, Printer } from 'lucide-react';
import toast from 'react-hot-toast';
import { LabelPrintModal } from '@/components/labels/label-print-modal';
import { useLabelTemplate } from '@/hooks/use-label-templates';

import { useSetting } from '@/hooks/use-settings';
import { usePriceGroups } from '@/hooks/use-price-groups';
import { useProductCategoryOptions } from '@/hooks/use-product-categories';
import { useCampaigns } from '@/hooks/use-events';
import { DEFAULT_CAT_LABELS } from '@/lib/utils/setting-defaults';

interface Props {
  productId?: string;
  mode?: 'view' | 'create' | 'duplicate';
  duplicateData?: { name: string; category: string; price: number; price_dealer: number; price_academy: number; price_purchase: number; price_groups?: Record<string, { price?: number | null; display_name?: string | null }> | null; description: string; imweb_product_no: string; barcode?: string; supplier_id: string };
  onClose: () => void;
  onCreated?: (id: string) => void;
}

export function ProductDetailPanel({ productId, mode = 'view', duplicateData, onClose, onCreated }: Props) {
  const router = useRouter();
  const queryClient = useQueryClient();
  const isCreateMode = mode === 'create' || mode === 'duplicate';
  const { data, isLoading } = useProduct(productId || '');
  const updateProduct = useUpdateProduct();
  const createProduct = useCreateProduct();
  const { data: sourcingData } = useSourcingByProduct(productId || '');
  const sourcing = sourcingData?.sourcing ?? null;
  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const priceGroups = usePriceGroups();
  const CATEGORY_LABEL = catLabels;
  const CATEGORY_OPTIONS = useProductCategoryOptions(); // 설정 SSOT + 시스템 고정(EVENT) 보장
  const productLabel = useLabelTemplate('product_40x20');
  const [editing, setEditing] = useState(false);
  const [showSerialModal, setShowSerialModal] = useState(false);
  const [showLabelModal, setShowLabelModal] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [showDeactivateConfirm, setShowDeactivateConfirm] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [form, setForm] = useState({
    sku: '', name: '', category: 'BL', price: 0, price_purchase: 0,
    price_group_values: {} as Record<string, { price: number; display_name: string }>,
    use_stock: true,
    description: '', imweb_product_no: '', barcode: '', supplier_id: '', product_group: '',
    purchase_name: '',
  });
  // EVENT 전용 tags (category='EVENT'일 때만 저장) — 폼 setForm 충돌 피하려 별도 state
  const [eventTags, setEventTags] = useState({ campaign_id: '', hand: 'right', event_type: 'blunt', spec: '' });
  const { data: eventCampaigns } = useCampaigns();

  // SKU 자동 채번
  async function fetchNextSku(cat: string) {
    try {
      const res = await fetch(`/api/products/next-sku?category=${cat}`);
      const data = await res.json();
      if (data.sku) setForm((prev) => ({ ...prev, sku: data.sku }));
    } catch { /* ignore */ }
  }

  useEffect(() => {
    setEditing(false);
    setShowSerialModal(false);
  }, [productId]);

  // create/duplicate 모드 초기화
  useEffect(() => {
    if (mode === 'create') {
      setForm({ sku: '', name: '', category: 'BL', price: 0, price_purchase: 0, price_group_values: {}, use_stock: true, description: '', imweb_product_no: '', barcode: '', supplier_id: '', product_group: '', purchase_name: '' });
      fetchNextSku('BL');
    } else if (mode === 'duplicate' && duplicateData) {
      // 기존 price_groups 또는 레거시 컬럼에서 price_group_values 복원
      const pgv: Record<string, { price: number; display_name: string }> = {};
      if (duplicateData.price_groups) {
        for (const [k, v] of Object.entries(duplicateData.price_groups)) {
          pgv[k] = { price: v?.price || 0, display_name: v?.display_name || '' };
        }
      } else {
        if (duplicateData.price_dealer > 0) pgv['dealer'] = { price: duplicateData.price_dealer, display_name: '' };
        if (duplicateData.price_academy > 0) pgv['academy'] = { price: duplicateData.price_academy, display_name: '' };
      }
      setForm({ sku: '', barcode: '', product_group: '', purchase_name: '', use_stock: true, name: '', category: duplicateData.category, price: duplicateData.price, price_purchase: duplicateData.price_purchase, price_group_values: pgv, description: duplicateData.description, imweb_product_no: duplicateData.imweb_product_no, supplier_id: duplicateData.supplier_id });
      fetchNextSku(duplicateData.category || 'BL');
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, duplicateData]);

  useEffect(() => {
    if (data?.product && !editing && mode === 'view') {
      const p = data.product;
      // price_groups JSONB → 폼 state 변환
      const pgv: Record<string, { price: number; display_name: string }> = {};
      if (p.price_groups) {
        for (const [k, v] of Object.entries(p.price_groups)) {
          pgv[k] = { price: v?.price || 0, display_name: v?.display_name || '' };
        }
      } else {
        // fallback: 레거시 컬럼에서 복원
        if (p.price_dealer > 0) pgv['dealer'] = { price: p.price_dealer, display_name: (p as Record<string, unknown>).dealer_name as string || '' };
        if (p.price_academy > 0) pgv['academy'] = { price: p.price_academy, display_name: (p as Record<string, unknown>).academy_name as string || '' };
      }
      setForm({
        sku: p.sku || '', name: p.name, category: p.category, price: p.price,
        price_purchase: p.price_purchase || 0,
        price_group_values: pgv,
        use_stock: p.stock_quantity !== -1,
        description: p.description || '', imweb_product_no: p.imweb_product_no || '',
        barcode: p.barcode || '', supplier_id: p.supplier_id || '', product_group: p.product_group || '',
        purchase_name: (p as Record<string, unknown>).purchase_name as string || '',
      });
      const t = (p.tags || {}) as Record<string, string>;
      setEventTags({ campaign_id: t.campaign_id || '', hand: t.hand || 'right', event_type: t.event_type || 'blunt', spec: t.spec || '' });
    }
  }, [data, editing, mode]);

  // 카테고리 변경 시 SKU 자동 재채번 (create/duplicate 모드)
  function handleCategoryChange(newCat: string) {
    setForm((prev) => ({ ...prev, category: newCat }));
    if (isCreateMode) fetchNextSku(newCat);
  }

  async function handleSave() {
    if (!productId) return;
    // 재고 미사용 전환: use_stock OFF → stock_quantity = -1
    const stockUpdate = !form.use_stock && data?.product?.stock_quantity !== -1
      ? { stock_quantity: -1 }
      : form.use_stock && data?.product?.stock_quantity === -1
        ? { stock_quantity: 0 }
        : {};

    // price_groups JSONB 조립
    const pgPayload: Record<string, { price?: number; display_name?: string }> = {};
    for (const [key, val] of Object.entries(form.price_group_values)) {
      if (val.price > 0 || val.display_name) {
        pgPayload[key] = { price: val.price || undefined, display_name: val.display_name || undefined };
      }
    }

    await updateProduct.mutateAsync({
      id: productId,
      name: form.name, category: form.category, price: form.price,
      price_dealer: form.price_group_values['dealer']?.price || 0,  // dual-write
      price_academy: form.price_group_values['academy']?.price || 0, // dual-write
      dealer_name: form.price_group_values['dealer']?.display_name || null, // dual-write
      academy_name: form.price_group_values['academy']?.display_name || null, // dual-write
      price_purchase: form.price_purchase,
      price_groups: pgPayload,
      ...stockUpdate,
      description: form.description || null, imweb_product_no: form.imweb_product_no || null,
      barcode: form.barcode || null, supplier_id: form.supplier_id || null,
      product_group: form.product_group || null,
      purchase_name: form.purchase_name || null,
      ...(form.category === 'EVENT'
        ? { tags: { campaign_id: eventTags.campaign_id || null, hand: eventTags.hand, event_type: eventTags.event_type, spec: eventTags.spec.trim() } }
        : {}),
    });
    setEditing(false);
  }

  async function handleCreate() {
    if (!form.sku.trim() || !form.name.trim()) { toast.error('SKU와 제품명을 입력해주세요'); return; }
    const pgPayload: Record<string, { price?: number; display_name?: string }> = {};
    for (const [key, val] of Object.entries(form.price_group_values)) {
      if (val.price > 0 || val.display_name) {
        pgPayload[key] = { price: val.price || undefined, display_name: val.display_name || undefined };
      }
    }
    const result = await createProduct.mutateAsync({
      sku: form.sku.trim(), name: form.name.trim(), category: form.category, price: form.price,
      price_dealer: form.price_group_values['dealer']?.price || undefined, // dual-write
      price_academy: form.price_group_values['academy']?.price || undefined, // dual-write
      price_purchase: form.price_purchase || undefined, price_groups: pgPayload,
      description: form.description.trim() || undefined,
      imweb_product_no: form.imweb_product_no.trim() || undefined, barcode: form.barcode.trim() || undefined,
      supplier_id: form.supplier_id || undefined,
      purchase_name: form.purchase_name.trim() || undefined,
      ...(form.category === 'EVENT'
        ? { tags: { campaign_id: eventTags.campaign_id || null, hand: eventTags.hand, event_type: eventTags.event_type, spec: eventTags.spec.trim() } }
        : {}),
    });
    if (result?.product?.id && onCreated) onCreated(result.product.id);
  }

  // EVENT 전용 입력 (캠페인/손/종류/옵션) — category='EVENT'일 때만. 등록·수정 폼 공통
  const eventInputsBlock = (
    <div className="rounded-lg border border-indigo-100 bg-indigo-50/50 p-3 space-y-3">
      <p className="text-[11px] text-indigo-700">고객 접수폼 분류값 — 이 캠페인 접수폼에만 노출됩니다. (재고 ≤3이면 마감임박 자동)</p>
      <div className="grid grid-cols-2 gap-3">
        <div>
          <label className="text-xs text-neutral-500">캠페인 *</label>
          <select value={eventTags.campaign_id} onChange={(e) => setEventTags({ ...eventTags, campaign_id: e.target.value })}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-stone-400">
            <option value="">캠페인 선택…</option>
            {(eventCampaigns || []).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </div>
        <div>
          <label className="text-xs text-neutral-500">손</label>
          <select value={eventTags.hand} onChange={(e) => setEventTags({ ...eventTags, hand: e.target.value })}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-stone-400">
            <option value="right">오른손</option>
            <option value="left">왼손</option>
          </select>
        </div>
        <div>
          <label className="text-xs text-neutral-500">종류</label>
          <select value={eventTags.event_type} onChange={(e) => setEventTags({ ...eventTags, event_type: e.target.value })}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-stone-400">
            <option value="blunt">블런트</option>
            <option value="thinning">틴닝</option>
            <option value="long">장가위</option>
            <option value="dry">DRY</option>
          </select>
        </div>
        <div>
          <label className="text-xs text-neutral-500">옵션(spec)</label>
          <input type="text" value={eventTags.spec} onChange={(e) => setEventTags({ ...eventTags, spec: e.target.value })}
            placeholder={eventTags.event_type === 'thinning' ? '예: 20-30%' : eventTags.event_type === 'dry' ? 'stroke/slicing/curve' : '예: 5.5인치'}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
        </div>
      </div>
    </div>
  );

  // create/duplicate 모드 — 등록 폼 렌더링
  if (isCreateMode) {
    return (
      <div className="flex-1 overflow-y-auto space-y-4 p-4">
        <div className="flex items-center justify-between">
          <h3 className="text-sm font-bold text-stone-900">{mode === 'duplicate' ? '제품 복제' : '제품 등록'}</h3>
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center"><X size={16} /></button>
        </div>
        <Card>
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">SKU (자동)</label>
                <input type="text" value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })}
                  placeholder="자동 채번" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">카테고리</label>
                <select value={form.category} onChange={(e) => handleCategoryChange(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400">
                  {CATEGORY_OPTIONS.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
                </select>
              </div>
            </div>
            {form.category === 'EVENT' && eventInputsBlock}
            <div>
              <label className="text-xs text-neutral-500">제품명 *</label>
              <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="마모루 블런트 6.0" autoFocus
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
            </div>
            {/* 재고 관리 토글 */}
            <div className="flex items-center justify-between py-1">
              <div>
                <span className="text-xs text-neutral-500">재고 관리</span>
                <p className="text-[10px] text-neutral-400">OFF 시 재고 수량을 추적하지 않습니다</p>
              </div>
              <button onClick={() => setForm({ ...form, use_stock: !form.use_stock })}
                className={`relative w-10 h-5 rounded-full transition ${form.use_stock ? 'bg-neutral-900' : 'bg-neutral-200'}`}>
                <span className={`absolute top-0.5 left-0.5 w-4 h-4 rounded-full bg-white shadow transition-transform ${form.use_stock ? 'translate-x-5' : ''}`} />
              </button>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">소매가</label>
                <input type="number" value={form.price || ''} onChange={(e) => setForm({ ...form, price: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
              {Object.entries(priceGroups).map(([key, def]) => (
                <div key={key}>
                  <label className="text-xs text-neutral-500">{def.label}</label>
                  <input type="number" value={form.price_group_values[key]?.price || ''} onChange={(e) => setForm({ ...form, price_group_values: { ...form.price_group_values, [key]: { ...form.price_group_values[key], price: parseInt(e.target.value) || 0, display_name: form.price_group_values[key]?.display_name || '' } } })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                </div>
              ))}
              <div>
                <label className="text-xs text-neutral-500">매입가</label>
                <input type="number" value={form.price_purchase || ''} onChange={(e) => setForm({ ...form, price_purchase: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
            </div>
            {/* 단가 그룹별 납품명 */}
            {Object.keys(priceGroups).length > 0 && (
              <div className="grid grid-cols-2 gap-3">
                {Object.entries(priceGroups).map(([key, def]) => (
                  <div key={`dn-${key}`}>
                    <label className="text-xs text-neutral-500">{def.label} 납품명</label>
                    <input type="text" value={form.price_group_values[key]?.display_name || ''} onChange={(e) => setForm({ ...form, price_group_values: { ...form.price_group_values, [key]: { ...form.price_group_values[key], price: form.price_group_values[key]?.price || 0, display_name: e.target.value } } })}
                      placeholder="미입력 시 기본 제품명"
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  </div>
                ))}
              </div>
            )}
            <div>
              <label className="text-xs text-neutral-500">매입처</label>
              <SupplierSelect value={form.supplier_id} onChange={(id) => setForm({ ...form, supplier_id: id })} placeholder="매입처 선택" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">아임웹 상품번호</label>
              <input type="text" value={form.imweb_product_no} onChange={(e) => setForm({ ...form, imweb_product_no: e.target.value })}
                placeholder="아임웹 상품관리에서 확인"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">제품군 (리뷰 그룹)</label>
              <input type="text" value={form.product_group} onChange={(e) => setForm({ ...form, product_group: e.target.value })}
                placeholder="예: R4, M5, CS600"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">발주명 (매입처 주문 시)</label>
              <input type="text" value={form.purchase_name} onChange={(e) => setForm({ ...form, purchase_name: e.target.value })}
                placeholder="매입처에서 사용하는 이름"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">설명</label>
              <textarea value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} rows={2}
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400 resize-none" />
            </div>
          </div>
        </Card>
        <div className="flex gap-2">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button className="flex-1" disabled={!form.sku.trim() || !form.name.trim() || createProduct.isPending} onClick={handleCreate}>
            {createProduct.isPending ? '등록 중...' : '제품 등록'}
          </Button>
        </div>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="space-y-4">
        <Skeleton className="h-8 w-48" />
        <Skeleton className="h-48 w-full" />
      </div>
    );
  }

  if (!data?.product) {
    return (
      <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
        <Package size={32} className="mb-2 opacity-50" />
        제품을 찾을 수 없습니다
      </div>
    );
  }

  const { product: p, supplier } = data;

  return (
    <div className="flex-1 overflow-y-auto space-y-4 pr-1">
      {/* 헤더 */}
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-2 min-w-0">
          <Package size={18} className="text-neutral-700 shrink-0" />
          <h3 className="text-sm font-bold text-stone-900 truncate">{p.name}</h3>
          <Badge className="bg-neutral-100 text-neutral-600 shrink-0">{p.sku}</Badge>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          {!editing ? (
            <>
              <Button
                variant="ghost"
                size="sm"
                onClick={() => setShowDeactivateConfirm(true)}
                title={p.is_active ? '비활성화' : '활성화'}
                className={p.is_active ? 'text-neutral-400' : 'text-red-500'}
              >
                {p.is_active ? <EyeOff size={14} /> : <Eye size={14} />}
              </Button>
              <Button
                variant="ghost"
                size="sm"
                onClick={() => setShowDeleteConfirm(true)}
                title="삭제"
                className="text-neutral-400 hover:text-red-500"
              >
                <Trash2 size={14} />
              </Button>
              <Button variant="ghost" size="sm" onClick={() => {
                if (onCreated && data?.product) {
                  onCreated('__duplicate__');
                }
              }} title="복제">
                <Copy size={14} />
              </Button>
              <Button variant="ghost" size="sm" onClick={() => setEditing(true)}>수정</Button>
            </>
          ) : (
            <>
              <Button variant="ghost" size="sm" onClick={() => setEditing(false)}>취소</Button>
              <Button size="sm" onClick={handleSave} disabled={updateProduct.isPending}>
                <Save size={14} />
                {updateProduct.isPending ? '저장 중...' : '저장'}
              </Button>
            </>
          )}
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center">
            <X size={16} />
          </button>
        </div>
      </div>

      {editing ? (
        <Card>
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">제품명</label>
                <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">카테고리</label>
                <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400">
                  {CATEGORY_OPTIONS.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
                </select>
              </div>
            </div>
            {form.category === 'EVENT' && eventInputsBlock}
            {/* 재고 관리 토글 */}
            <div className="flex items-center justify-between py-1">
              <div>
                <span className="text-xs text-neutral-500">재고 관리</span>
                <p className="text-[10px] text-neutral-400">OFF 시 재고 수량을 추적하지 않습니다</p>
              </div>
              <button onClick={() => setForm({ ...form, use_stock: !form.use_stock })}
                className={`relative w-10 h-5 rounded-full transition ${form.use_stock ? 'bg-neutral-900' : 'bg-neutral-200'}`}>
                <span className={`absolute top-0.5 left-0.5 w-4 h-4 rounded-full bg-white shadow transition-transform ${form.use_stock ? 'translate-x-5' : ''}`} />
              </button>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">소매가</label>
                <input type="number" value={form.price || ''} onChange={(e) => setForm({ ...form, price: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
              {Object.entries(priceGroups).map(([key, def]) => (
                <div key={key}>
                  <label className="text-xs text-neutral-500">{def.label}</label>
                  <input type="number" value={form.price_group_values[key]?.price || ''} onChange={(e) => setForm({ ...form, price_group_values: { ...form.price_group_values, [key]: { ...form.price_group_values[key], price: parseInt(e.target.value) || 0, display_name: form.price_group_values[key]?.display_name || '' } } })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                </div>
              ))}
              <div>
                <label className="text-xs text-neutral-500">매입가</label>
                <input type="number" value={form.price_purchase || ''} onChange={(e) => setForm({ ...form, price_purchase: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
            </div>
            {/* 단가 그룹별 납품명 */}
            {Object.keys(priceGroups).length > 0 && (
              <div className="grid grid-cols-2 gap-3">
                {Object.entries(priceGroups).map(([key, def]) => (
                  <div key={`dn-${key}`}>
                    <label className="text-xs text-neutral-500">{def.label} 납품명</label>
                    <input type="text" value={form.price_group_values[key]?.display_name || ''} onChange={(e) => setForm({ ...form, price_group_values: { ...form.price_group_values, [key]: { ...form.price_group_values[key], price: form.price_group_values[key]?.price || 0, display_name: e.target.value } } })}
                      placeholder="미입력 시 기본 제품명"
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  </div>
                ))}
              </div>
            )}
            <div>
              <label className="text-xs text-neutral-500">매입처</label>
              <SupplierSelect value={form.supplier_id} displayName={supplier?.name}
                onChange={(id) => setForm({ ...form, supplier_id: id })} placeholder="매입처 선택" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">제품군 (리뷰 그룹)</label>
              <input type="text" value={form.product_group} onChange={(e) => setForm({ ...form, product_group: e.target.value })}
                placeholder="예: R4, M5, CS600"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">발주명 (매입처 주문 시)</label>
              <input type="text" value={form.purchase_name} onChange={(e) => setForm({ ...form, purchase_name: e.target.value })}
                placeholder="매입처에서 사용하는 이름"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">설명</label>
              <textarea value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} rows={2}
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400 resize-none" />
            </div>
          </div>
        </Card>
      ) : (
        <>
          {/* 이미지 */}
          {p.image_url && (
            <div className="rounded-lg overflow-hidden bg-neutral-50">
              <img src={p.image_url} alt={p.name} className="w-full h-40 object-contain" />
            </div>
          )}

          {/* 기본 정보 */}
          <Card>
            <div className="grid grid-cols-3 gap-3 text-sm">
              <div>
                <span className="text-xs text-neutral-500">카테고리</span>
                <p>{CATEGORY_LABEL[p.category] || p.category}</p>
              </div>
              <div>
                <span className="text-xs text-neutral-500">재고</span>
                {p.stock_quantity === -1 ? (
                  <p><Badge className="bg-neutral-100 text-neutral-500 text-[10px]">미사용</Badge></p>
                ) : (
                  <p className={p.stock_quantity > 0 ? 'font-bold' : 'font-bold text-red-500'}>{p.stock_quantity}개</p>
                )}
              </div>
              <div>
                <span className="text-xs text-neutral-500">상태</span>
                <p>
                  <Badge className={p.is_active ? 'bg-green-100 text-green-700' : 'bg-neutral-100 text-neutral-500'}>
                    {p.is_active ? '활성' : '비활성'}
                  </Badge>
                </p>
              </div>
            </div>
          </Card>

          {/* 가격 */}
          <Card>
            <h4 className="text-xs text-neutral-500 mb-2">가격 정보</h4>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <p className="text-xs text-neutral-400">소매가</p>
                <p className="text-sm font-bold">{formatKRW(p.price)}</p>
              </div>
              {Object.entries(priceGroups).map(([key, def]) => {
                const groupPrice = p.price_groups?.[key]?.price;
                return (
                  <div key={key}>
                    <p className="text-xs text-neutral-400">{def.label}</p>
                    <p className={`text-sm font-bold text-${def.color}-600`}>{groupPrice && groupPrice > 0 ? formatKRW(groupPrice) : '-'}</p>
                  </div>
                );
              })}
              <div>
                <p className="text-xs text-neutral-400">매입가</p>
                <p className="text-sm font-bold text-neutral-600">{p.price_purchase > 0 ? formatKRW(p.price_purchase) : '-'}</p>
              </div>
            </div>
          </Card>

          {/* 부가 정보 */}
          <Card>
            {supplier && (
              <div className="mb-3">
                <span className="text-xs text-neutral-500">매입처</span>
                <p className="text-sm">{supplier.name}</p>
              </div>
            )}
            {p.product_group && (
              <div className="mb-3">
                <span className="text-xs text-neutral-500">제품군</span>
                <p className="text-sm"><Badge className="bg-blue-50 text-blue-700">{p.product_group}</Badge></p>
              </div>
            )}
            {p.imweb_product_no && (
              <div className="mb-3">
                <span className="text-xs text-neutral-500">아임웹</span>
                <p className="text-sm font-mono">#{p.imweb_product_no}</p>
              </div>
            )}
            {p.description && (
              <div>
                <span className="text-xs text-neutral-500">설명</span>
                <p className="text-sm text-neutral-600 whitespace-pre-wrap">{p.description}</p>
              </div>
            )}
          </Card>

          {/* 소싱 정보 (1688) — 연결된 소싱 품목 있을 때만 */}
          {sourcing && (
            <Card>
              <h4 className="text-xs text-neutral-500 mb-2">🏭 소싱 정보 (1688)</h4>
              <div className="space-y-2 text-sm">
                {sourcing.supplier_name && (
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs text-neutral-500 w-14 flex-shrink-0">업체</span>
                    <span className="text-neutral-800 truncate">{sourcing.supplier_name}</span>
                    {sourcing.supplier_url && (
                      <a href={sourcing.supplier_url} target="_blank" rel="noreferrer" className="text-blue-600 flex-shrink-0"><ExternalLink size={12} /></a>
                    )}
                  </div>
                )}
                {sourcing.vendor_url && (
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs text-neutral-500 w-14 flex-shrink-0">품목링크</span>
                    <a href={sourcing.vendor_url} target="_blank" rel="noreferrer" className="text-blue-600 inline-flex items-center gap-0.5 text-xs truncate">1688 상품 <ExternalLink size={11} /></a>
                  </div>
                )}
                <div className="flex items-center gap-1.5">
                  <span className="text-xs text-neutral-500 w-14 flex-shrink-0">소싱 단가</span>
                  <span className="text-neutral-800">¥{sourcing.unit_price} <span className="text-neutral-400">≈ ₩{sourcing.krw_price.toLocaleString()}</span></span>
                </div>
                {sourcing.po_number && (
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs text-neutral-500 w-14 flex-shrink-0">소싱회차</span>
                    <span className="font-mono text-xs text-neutral-500">{sourcing.po_number}</span>
                  </div>
                )}
              </div>
            </Card>
          )}

          {/* 퀵 액션 */}
          <Card>
            <h4 className="text-xs text-neutral-500 mb-2">빠른 작업</h4>
            <div className="space-y-2">
              <Button variant="primary" size="sm" onClick={() => setShowSerialModal(true)} className="w-full">
                <ArrowRightLeft size={14} />
                창고 이동
              </Button>
              {p.category !== 'SUP' && (
                <Button variant="secondary" size="sm" onClick={() => setShowLabelModal(true)} className="w-full">
                  <Printer size={14} />
                  바코드 라벨 출력
                </Button>
              )}
              <button
                onClick={() => router.push(`/products/${productId}/serials`)}
                className="w-full text-center text-xs text-neutral-400 hover:text-neutral-600 py-1 transition"
              >
                시리얼 상세 관리 →
              </button>
            </div>
          </Card>
        </>
      )}

      {/* 품목 바코드 라벨 출력 */}
      {showLabelModal && (
        <LabelPrintModal
          template={productLabel}
          data={{ product: p.name, sku: p.barcode || p.sku }}
          title="품목 바코드 라벨"
          onClose={() => setShowLabelModal(false)}
        />
      )}

      {/* 시리얼 빠른 등록 모달 */}
      <SerialQuickModal
        open={showSerialModal}
        onClose={() => setShowSerialModal(false)}
        productId={productId!}
        productSku={p.sku}
        rawStock={p.raw_stock || 0}
        onSuccess={() => {
          queryClient.invalidateQueries({ queryKey: ['product', productId] });
          queryClient.invalidateQueries({ queryKey: ['products'] });
          queryClient.invalidateQueries({ queryKey: ['inventory'] });
          queryClient.invalidateQueries({ queryKey: ['serials'] });
        }}
      />

      {/* 비활성화/활성화 확인 모달 */}
      <ConfirmModal
        open={showDeactivateConfirm}
        onClose={() => setShowDeactivateConfirm(false)}
        onConfirm={async () => {
          if (!productId) return;
          await updateProduct.mutateAsync({ id: productId, is_active: !p.is_active });
          queryClient.invalidateQueries({ queryKey: ['product', productId] });
          queryClient.invalidateQueries({ queryKey: ['products'] });
          queryClient.invalidateQueries({ queryKey: ['inventory'] });
          toast.success(p.is_active ? '비활성화됨' : '활성화됨');
        }}
        title={p.is_active ? '제품 비활성화' : '제품 활성화'}
        message={p.is_active
          ? `${p.name}을(를) 비활성화합니다. 목록에서 숨겨지며 판매 시 표시되지 않습니다.`
          : `${p.name}을(를) 다시 활성화합니다.`}
        confirmLabel={p.is_active ? '비활성화' : '활성화'}
        variant={p.is_active ? 'danger' : 'default'}
      />

      {/* 삭제 확인 모달 */}
      {showDeleteConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-xl shadow-xl p-6 w-full max-w-sm mx-4">
            <h3 className="text-lg font-bold text-neutral-900">제품 삭제</h3>
            <p className="text-sm text-neutral-600 mt-2">
              <span className="font-semibold">{p.name}</span>을(를) 삭제하시겠습니까?
            </p>
            <p className="text-xs text-neutral-400 mt-1">
              시리얼·판매·계약서가 연결된 제품은 삭제할 수 없습니다.
            </p>
            <div className="flex gap-2 mt-5 justify-end">
              <Button variant="ghost" size="sm" onClick={() => setShowDeleteConfirm(false)}>취소</Button>
              <Button
                variant="danger"
                size="sm"
                loading={deleting}
                onClick={async () => {
                  setDeleting(true);
                  try {
                    const res = await fetch(`/api/products/${productId}`, { method: 'DELETE' });
                    const data = await res.json();
                    if (!res.ok) { toast.error(data.error || '삭제 실패'); return; }
                    toast.success('제품이 삭제되었습니다');
                    queryClient.invalidateQueries({ queryKey: ['products'] });
                    queryClient.invalidateQueries({ queryKey: ['inventory'] });
                    onClose();
                  } catch (err) {
                    toast.error(String(err));
                  } finally {
                    setDeleting(false);
                    setShowDeleteConfirm(false);
                  }
                }}
              >
                삭제
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ── 시리얼 빠른 등록 모달 ── */

type Zone = 'raw' | 'ready' | 'display';
const ZONE_LABEL: Record<Zone, string> = { raw: '보관', ready: '준비', display: '디스플레이' };
const ZONE_COLOR: Record<Zone, string> = {
  raw: 'border-neutral-900 bg-neutral-900 text-white',
  ready: 'border-green-600 bg-green-600 text-white',
  display: 'border-blue-600 bg-blue-600 text-white',
};
const ZONE_INACTIVE = 'border-neutral-200 bg-white text-neutral-600 hover:border-neutral-300';

function SerialQuickModal({ open, onClose, productId, productSku, rawStock, onSuccess }: {
  open: boolean;
  onClose: () => void;
  productId: string;
  productSku: string;
  rawStock: number;
  onSuccess: () => void;
}) {
  const [from, setFrom] = useState<Zone>('raw');
  const [to, setTo] = useState<Zone>('ready');
  const [count, setCount] = useState(1);
  const [nextSerial, setNextSerial] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [loadingNext, setLoadingNext] = useState(false);

  // 보관→시리얼 창고 이동 시에만 시리얼 생성 필요
  const needsSerial = from === 'raw';
  // 시리얼 창고→보관 역이동 (시리얼 삭제 + raw_stock 복원)
  const isReverse = to === 'raw' && from !== 'raw';
  // 시리얼 창고 간 이동 (zone 변경만)
  const isZoneTransfer = from !== 'raw' && to !== 'raw';

  // 모달 열릴 때
  useEffect(() => {
    if (open) {
      setFrom('raw');
      setTo('ready');
      setCount(1);
      setLoadingNext(true);
      fetch('/api/serials/batch')
        .then((res) => res.json())
        .then((data) => { if (data.next_serial) setNextSerial(data.next_serial); })
        .catch(() => setNextSerial(''))
        .finally(() => setLoadingNext(false));
    }
  }, [open]);

  // 출발지 변경 시 도착지 자동 조정
  function handleFromChange(zone: Zone) {
    setFrom(zone);
    if (zone === to) setTo(zone === 'raw' ? 'ready' : 'raw');
  }

  // 도착지에서 출발지와 같은 건 제외
  const availableTo = (['raw', 'ready', 'display'] as Zone[]).filter((z) => z !== from);

  async function handleSubmit() {
    if (count < 1) return;
    setSubmitting(true);
    try {
      if (needsSerial) {
        // 보관 → 시리얼 창고: batch API 호출 (서버에서 다음 M{YY}-{NNNN} 자동 채번)
        const res = await fetch('/api/serials/batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ product_id: productId, count, warehouse_zone: to }),
        });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: '생성 실패' }));
          toast.error(errData.error || '생성 실패'); setSubmitting(false); return;
        }
        const data = await res.json();
        toast.success(`${data.created || count}개 → ${ZONE_LABEL[to]} 창고 (시리얼 생성)`);
      } else if (isZoneTransfer) {
        // 시리얼 창고 간 이동: zone 변경
        // 해당 제품의 from zone 시리얼을 count개만 변경
        const res = await fetch(`/api/serials/move`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ product_id: productId, from_zone: from, to_zone: to, count }),
        });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: '이동 실패' }));
          toast.error(errData.error || '이동 실패'); setSubmitting(false); return;
        }
        toast.success(`${count}개 ${ZONE_LABEL[from]} → ${ZONE_LABEL[to]} 이동`);
      } else if (isReverse) {
        // 시리얼→보관 역이동: 시리얼 삭제 + raw_stock 증가
        const res = await fetch(`/api/serials/move`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ product_id: productId, from_zone: from, to_zone: 'raw', count }),
        });
        if (!res.ok) {
          const errData = await res.json().catch(() => ({ error: '이동 실패' }));
          toast.error(errData.error || '이동 실패'); setSubmitting(false); return;
        }
        toast.success(`${count}개 ${ZONE_LABEL[from]} → 보관 (시리얼 해제)`);
      }
      onSuccess();
      onClose();
    } catch (err) {
      toast.error(`오류: ${String(err)}`);
    } finally {
      setSubmitting(false);
    }
  }

  if (!open) return null;

  return (
    <Modal open={open} onClose={onClose} title="창고 이동" className="max-w-md">
      <div className="space-y-4">
        {/* 현재 재고 현황 */}
        <div className="flex items-center gap-3 px-3 py-2 rounded-lg bg-neutral-50 text-xs text-neutral-500">
          <span>보관 <strong className="text-neutral-800">{rawStock}</strong></span>
          <span className="text-neutral-300">|</span>
          <span>총 재고는 창고·재고 페이지에서 확인</span>
        </div>

        {/* 출발지 → 도착지 */}
        <div className="flex items-center gap-3">
          {/* 출발지 */}
          <div className="flex-1">
            <label className="text-xs text-neutral-500 mb-1.5 block">출발지</label>
            <div className="flex gap-1.5">
              {(['raw', 'ready', 'display'] as Zone[]).map((z) => (
                <button key={z} onClick={() => handleFromChange(z)}
                  className={`flex-1 py-2 rounded-lg border-2 text-xs font-semibold transition ${from === z ? ZONE_COLOR[z] : ZONE_INACTIVE}`}>
                  {ZONE_LABEL[z]}
                </button>
              ))}
            </div>
          </div>
          <ArrowRight size={16} className="text-neutral-300 mt-5 shrink-0" />
          {/* 도착지 */}
          <div className="flex-1">
            <label className="text-xs text-neutral-500 mb-1.5 block">도착지</label>
            <div className="flex gap-1.5">
              {availableTo.map((z) => (
                <button key={z} onClick={() => setTo(z)}
                  className={`flex-1 py-2 rounded-lg border-2 text-xs font-semibold transition ${to === z ? ZONE_COLOR[z] : ZONE_INACTIVE}`}>
                  {ZONE_LABEL[z]}
                </button>
              ))}
            </div>
          </div>
        </div>

        {/* 수량 */}
        <div>
          <label className="text-xs text-neutral-500 mb-1 block">이동 수량</label>
          <input type="number" value={count}
            onChange={(e) => setCount(Math.max(1, parseInt(e.target.value) || 1))}
            min={1}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
        </div>

        {/* 시리얼 생성 UI — 보관→시리얼 창고 일 때만 (M{YY}-{NNNN} 자동 채번) */}
        {needsSerial && (
          <div className="border border-neutral-200 rounded-lg p-3 space-y-1">
            <label className="text-xs font-semibold text-neutral-600">시리얼 자동생성</label>
            <p className="text-xs text-neutral-500">
              {loadingNext ? '다음 번호 조회 중...' : (
                <>다음 번호 <span className="font-mono font-semibold text-neutral-700">{nextSerial || '-'}</span> 부터 <span className="font-semibold">{count}</span>개 자동 생성</>
              )}
            </p>
          </div>
        )}

        {/* 미리보기 */}
        <div className="text-xs bg-neutral-50 rounded-lg p-3">
          <p className="text-neutral-600 font-medium">
            {ZONE_LABEL[from]} → {ZONE_LABEL[to]} · {count}개
            {needsSerial && ' (시리얼 생성)'}
            {isReverse && ' (시리얼 해제)'}
            {isZoneTransfer && ' (zone 변경)'}
          </p>
        </div>

        {/* 버튼 */}
        <div className="flex gap-2">
          <Button variant="ghost" size="sm" onClick={onClose} className="flex-1">취소</Button>
          <Button size="sm" onClick={handleSubmit}
            disabled={count < 1 || (needsSerial && rawStock < count) || submitting}
            loading={submitting} className="flex-1">
            {submitting ? '처리 중...' : `${count}개 이동`}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
