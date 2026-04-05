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
import { useProduct, useUpdateProduct, useCreateProduct } from '@/hooks/use-product-detail';
import { formatKRW } from '@/lib/utils/format';
import { Save, Package, Hash, X, Receipt, Boxes, Plus, Archive, Copy, Eye, EyeOff, Trash2 } from 'lucide-react';
import toast from 'react-hot-toast';

import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CAT_LABELS } from '@/lib/utils/setting-defaults';

interface Props {
  productId?: string;
  mode?: 'view' | 'create' | 'duplicate';
  duplicateData?: { name: string; category: string; price: number; price_dealer: number; price_academy: number; price_purchase: number; description: string; imweb_product_no: string; barcode?: string; supplier_id: string };
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
  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const categories = useSetting<string[]>('inventory.categories', Object.keys(DEFAULT_CAT_LABELS));
  const CATEGORY_LABEL = catLabels;
  const CATEGORY_OPTIONS = categories.map((c) => ({ value: c, label: catLabels[c] || c }));
  const [editing, setEditing] = useState(false);
  const [showSerialModal, setShowSerialModal] = useState(false);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [showDeactivateConfirm, setShowDeactivateConfirm] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [form, setForm] = useState({
    sku: '', name: '', category: 'BL', price: 0, price_dealer: 0, price_academy: 0, price_purchase: 0,
    dealer_name: '', academy_name: '',
    description: '', imweb_product_no: '', barcode: '', supplier_id: '', product_group: '',
  });

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
      setForm({ sku: '', name: '', category: 'BL', price: 0, price_dealer: 0, price_academy: 0, price_purchase: 0, dealer_name: '', academy_name: '', description: '', imweb_product_no: '', barcode: '', supplier_id: '', product_group: '' });
      fetchNextSku('BL');
    } else if (mode === 'duplicate' && duplicateData) {
      setForm({ sku: '', barcode: '', product_group: '', dealer_name: '', academy_name: '', ...duplicateData, name: '' });
      fetchNextSku(duplicateData.category || 'BL');
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, duplicateData]);

  useEffect(() => {
    if (data?.product && !editing && mode === 'view') {
      const p = data.product;
      setForm({
        sku: p.sku || '', name: p.name, category: p.category, price: p.price,
        price_dealer: p.price_dealer || 0, price_academy: p.price_academy || 0, price_purchase: p.price_purchase || 0,
        dealer_name: (p as Record<string, unknown>).dealer_name as string || '', academy_name: (p as Record<string, unknown>).academy_name as string || '',
        description: p.description || '', imweb_product_no: p.imweb_product_no || '',
        barcode: p.barcode || '', supplier_id: p.supplier_id || '', product_group: p.product_group || '',
      });
    }
  }, [data, editing, mode]);

  // 카테고리 변경 시 SKU 자동 재채번 (create/duplicate 모드)
  function handleCategoryChange(newCat: string) {
    setForm((prev) => ({ ...prev, category: newCat }));
    if (isCreateMode) fetchNextSku(newCat);
  }

  async function handleSave() {
    if (!productId) return;
    await updateProduct.mutateAsync({
      id: productId,
      name: form.name, category: form.category, price: form.price,
      price_dealer: form.price_dealer, price_academy: form.price_academy, price_purchase: form.price_purchase,
      dealer_name: form.dealer_name || null, academy_name: form.academy_name || null,
      description: form.description || null, imweb_product_no: form.imweb_product_no || null,
      barcode: form.barcode || null, supplier_id: form.supplier_id || null,
      product_group: form.product_group || null,
    });
    setEditing(false);
  }

  async function handleCreate() {
    if (!form.sku.trim() || !form.name.trim()) { toast.error('SKU와 제품명을 입력해주세요'); return; }
    const result = await createProduct.mutateAsync({
      sku: form.sku.trim(), name: form.name.trim(), category: form.category, price: form.price,
      price_dealer: form.price_dealer || undefined, price_academy: form.price_academy || undefined,
      price_purchase: form.price_purchase || undefined, description: form.description.trim() || undefined,
      imweb_product_no: form.imweb_product_no.trim() || undefined, barcode: form.barcode.trim() || undefined,
      supplier_id: form.supplier_id || undefined,
    });
    if (result?.product?.id && onCreated) onCreated(result.product.id);
  }

  // create/duplicate 모드 — 등록 폼 렌더링
  if (isCreateMode) {
    return (
      <div className="flex-1 overflow-y-auto space-y-4 p-4">
        <div className="flex items-center justify-between">
          <h3 className="text-sm font-bold text-indigo-black">{mode === 'duplicate' ? '제품 복제' : '제품 등록'}</h3>
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center"><X size={16} /></button>
        </div>
        <Card>
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">SKU (자동)</label>
                <input type="text" value={form.sku} onChange={(e) => setForm({ ...form, sku: e.target.value })}
                  placeholder="자동 채번" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm font-mono focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">카테고리</label>
                <select value={form.category} onChange={(e) => handleCategoryChange(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40">
                  {CATEGORY_OPTIONS.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
                </select>
              </div>
            </div>
            <div>
              <label className="text-xs text-neutral-500">제품명 *</label>
              <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="마모루 블런트 6.0" autoFocus
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">소매가</label>
                <input type="number" value={form.price || ''} onChange={(e) => setForm({ ...form, price: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">딜러가</label>
                <input type="number" value={form.price_dealer || ''} onChange={(e) => setForm({ ...form, price_dealer: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">아카데미가</label>
                <input type="number" value={form.price_academy || ''} onChange={(e) => setForm({ ...form, price_academy: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">매입가</label>
                <input type="number" value={form.price_purchase || ''} onChange={(e) => setForm({ ...form, price_purchase: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
            </div>
            {/* B2B 납품명 */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">딜러 납품명</label>
                <input type="text" value={form.dealer_name} onChange={(e) => setForm({ ...form, dealer_name: e.target.value })}
                  placeholder="미입력 시 기본 제품명"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">아카데미 납품명</label>
                <input type="text" value={form.academy_name} onChange={(e) => setForm({ ...form, academy_name: e.target.value })}
                  placeholder="미입력 시 기본 제품명"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
            </div>
            <div>
              <label className="text-xs text-neutral-500">매입처</label>
              <SupplierSelect value={form.supplier_id} onChange={(id) => setForm({ ...form, supplier_id: id })} placeholder="매입처 선택" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">아임웹 상품번호</label>
              <input type="text" value={form.imweb_product_no} onChange={(e) => setForm({ ...form, imweb_product_no: e.target.value })}
                placeholder="아임웹 상품관리에서 확인"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">제품군 (리뷰 그룹)</label>
              <input type="text" value={form.product_group} onChange={(e) => setForm({ ...form, product_group: e.target.value })}
                placeholder="예: R4, M5, CS600"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">설명</label>
              <textarea value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} rows={2}
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none" />
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
          <h3 className="text-sm font-bold text-indigo-black truncate">{p.name}</h3>
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
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">카테고리</label>
                <select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40">
                  {CATEGORY_OPTIONS.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
                </select>
              </div>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">소매가</label>
                <input type="number" value={form.price || ''} onChange={(e) => setForm({ ...form, price: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">딜러가</label>
                <input type="number" value={form.price_dealer || ''} onChange={(e) => setForm({ ...form, price_dealer: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">아카데미가</label>
                <input type="number" value={form.price_academy || ''} onChange={(e) => setForm({ ...form, price_academy: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">매입가</label>
                <input type="number" value={form.price_purchase || ''} onChange={(e) => setForm({ ...form, price_purchase: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
            </div>
            {/* B2B 납품명 */}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">딜러 납품명</label>
                <input type="text" value={form.dealer_name} onChange={(e) => setForm({ ...form, dealer_name: e.target.value })}
                  placeholder="미입력 시 기본 제품명"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">아카데미 납품명</label>
                <input type="text" value={form.academy_name} onChange={(e) => setForm({ ...form, academy_name: e.target.value })}
                  placeholder="미입력 시 기본 제품명"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
            </div>
            <div>
              <label className="text-xs text-neutral-500">매입처</label>
              <SupplierSelect value={form.supplier_id} displayName={supplier?.name}
                onChange={(id) => setForm({ ...form, supplier_id: id })} placeholder="매입처 선택" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">제품군 (리뷰 그룹)</label>
              <input type="text" value={form.product_group} onChange={(e) => setForm({ ...form, product_group: e.target.value })}
                placeholder="예: R4, M5, CS600"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">설명</label>
              <textarea value={form.description} onChange={(e) => setForm({ ...form, description: e.target.value })} rows={2}
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none" />
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
              <div>
                <p className="text-xs text-neutral-400">딜러가</p>
                <p className="text-sm font-bold text-purple-600">{p.price_dealer > 0 ? formatKRW(p.price_dealer) : '-'}</p>
              </div>
              <div>
                <p className="text-xs text-neutral-400">아카데미가</p>
                <p className="text-sm font-bold text-emerald-600">{p.price_academy > 0 ? formatKRW(p.price_academy) : '-'}</p>
              </div>
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

          {/* 퀵 액션 */}
          <Card>
            <h4 className="text-xs text-neutral-500 mb-2">빠른 작업</h4>
            <div className="space-y-2">
              <Button variant="primary" size="sm" onClick={() => setShowSerialModal(true)} className="w-full">
                <Plus size={14} />
                시리얼 등록 · 창고배치
              </Button>
              <div className="flex gap-2">
                <Button variant="secondary" size="sm" onClick={() => router.push('/sales/new')} className="flex-1">
                  <Receipt size={14} />
                  판매 등록
                </Button>
                <Button variant="secondary" size="sm" onClick={() => router.push('/inventory')} className="flex-1">
                  <Boxes size={14} />
                  창고·재고
                </Button>
              </div>
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

const ZONE_OPTIONS: { key: 'raw' | 'ready' | 'display'; label: string; desc: string; color: string; activeColor: string }[] = [
  { key: 'raw', label: '보관', desc: '매입 원본 보관', color: 'border-neutral-200 bg-white text-neutral-600', activeColor: 'border-neutral-900 bg-neutral-900 text-white' },
  { key: 'ready', label: '준비', desc: '마모루 각인 완료, B2C 출고 가능', color: 'border-neutral-200 bg-white text-neutral-600', activeColor: 'border-green-600 bg-green-600 text-white' },
  { key: 'display', label: '디스플레이', desc: '고객 전시 샘플 (가방)', color: 'border-neutral-200 bg-white text-neutral-600', activeColor: 'border-blue-600 bg-blue-600 text-white' },
];

function SerialQuickModal({ open, onClose, productId, productSku, rawStock, onSuccess }: {
  open: boolean;
  onClose: () => void;
  productId: string;
  productSku: string;
  rawStock: number;
  onSuccess: () => void;
}) {
  const [startNumber, setStartNumber] = useState('');
  const [count, setCount] = useState(1);
  const [zone, setZone] = useState<'raw' | 'ready' | 'display'>('raw');
  const [submitting, setSubmitting] = useState(false);
  const [loadingNext, setLoadingNext] = useState(false);

  // 모달 열릴 때 — 다음 시리얼번호 자동 조회
  useEffect(() => {
    if (open) {
      setCount(1);
      setZone('raw');
      setLoadingNext(true);
      fetch('/api/serials/batch')
        .then((res) => res.json())
        .then((data) => {
          if (data.next_start) setStartNumber(String(data.next_start));
        })
        .catch(() => setStartNumber(''))
        .finally(() => setLoadingNext(false));
    }
  }, [open]);

  const zoneLabel = { raw: '보관', ready: '준비', display: '디스플레이' }[zone];

  async function handleSubmit() {
    if (!startNumber || count < 1) return;
    setSubmitting(true);
    try {
      const res = await fetch('/api/serials/batch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          product_id: productId,
          start_number: parseInt(startNumber),
          count,
          warehouse_zone: zone,
        }),
      });
      if (!res.ok) {
        const errData = await res.json().catch(() => ({ error: '알 수 없는 오류' }));
        toast.error(errData.error || '생성 실패');
        return;
      }
      const data = await res.json();
      toast.success(`${data.created || count}개 시리얼 → ${zoneLabel} 창고에 등록`);
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
    <Modal open={open} onClose={onClose} title="시리얼 등록 · 창고배치" className="max-w-md">
      <div className="space-y-4">
        {/* 보관 재고 표시 */}
        <div className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm ${rawStock > 0 ? 'bg-blue-50 text-blue-700' : 'bg-red-50 text-red-700'}`}>
          <Archive size={14} />
          보관창고: <strong>{rawStock}개</strong>
          {rawStock === 0 && <span className="text-xs">(재고 부족 — 입고 필요)</span>}
        </div>

        {/* 시작번호 + 수량 */}
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">시작번호 *</label>
            <input
              type="number"
              value={startNumber}
              onChange={(e) => setStartNumber(e.target.value)}
              placeholder={loadingNext ? '조회 중...' : '13790001'}
              disabled={loadingNext}
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm font-mono focus:outline-none focus:ring-2 focus:ring-terracotta/40 disabled:opacity-50"
            />
            <p className="text-[10px] text-neutral-400 mt-0.5">이전 번호 이어서 자동 입력</p>
          </div>
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">수량</label>
            <input
              type="number"
              value={count}
              onChange={(e) => setCount(Math.max(1, Math.min(rawStock || 100, parseInt(e.target.value) || 1)))}
              min={1}
              max={rawStock || 100}
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            />
          </div>
        </div>

        {/* 창고 선택 — 3개 */}
        <div>
          <label className="text-xs text-neutral-500 mb-2 block">등록할 창고</label>
          <div className="flex gap-2">
            {ZONE_OPTIONS.map((opt) => (
              <button
                key={opt.key}
                onClick={() => setZone(opt.key)}
                className={`flex-1 flex items-center justify-center gap-1.5 px-2 py-2.5 rounded-lg border-2 transition text-sm font-medium ${
                  zone === opt.key ? opt.activeColor : opt.color
                } hover:border-neutral-300`}
              >
                {opt.label}
              </button>
            ))}
          </div>
          <p className="text-[11px] text-neutral-400 mt-1.5">
            {ZONE_OPTIONS.find((o) => o.key === zone)?.desc}
          </p>
        </div>

        {/* 미리보기 */}
        {startNumber && count > 0 && (
          <div className="text-xs bg-neutral-50 rounded-lg p-3 space-y-1">
            <p className="text-neutral-500">
              생성 범위: <span className="font-mono font-semibold">{String(parseInt(startNumber) || 0).padStart(8, '0')}</span> ~ <span className="font-mono font-semibold">{String((parseInt(startNumber) || 0) + count - 1).padStart(8, '0')}</span>
            </p>
            <p className="text-neutral-400">
              {count}개 → {zoneLabel} 창고
            </p>
          </div>
        )}

        {/* 버튼 */}
        <div className="flex gap-2">
          <Button variant="ghost" size="sm" onClick={onClose} className="flex-1">
            취소
          </Button>
          <Button
            size="sm"
            onClick={handleSubmit}
            disabled={!startNumber || count < 1 || count > rawStock || rawStock === 0 || submitting}
            loading={submitting}
            className="flex-1"
          >
            {submitting ? '생성 중...' : `${count}개 등록`}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
