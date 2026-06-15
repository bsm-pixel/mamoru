'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useCreateProduct } from '@/hooks/use-product-detail';
import { SupplierSelect } from '@/components/ui/supplier-select';
import { usePriceGroups } from '@/hooks/use-price-groups';
import { ArrowLeft, CheckCircle, XCircle } from 'lucide-react';

const CATEGORY_OPTIONS = [
  { value: 'BL', label: '블런트' },
  { value: 'TH', label: '틴닝' },
  { value: 'LO', label: '장가위' },
  { value: 'SL', label: '슬라이싱' },
  { value: 'EVENT', label: 'EVENT (재고 전환 이벤트)' },
];

export default function NewProductPage() {
  const router = useRouter();
  const createProduct = useCreateProduct();
  const priceGroups = usePriceGroups();
  const [skuStatus, setSkuStatus] = useState<'idle' | 'checking' | 'available' | 'duplicate'>('idle');

  const [form, setForm] = useState({
    sku: '',
    name: '',
    category: 'BL',
    price: 0,
    price_purchase: 0,
    price_group_values: {} as Record<string, { price: number; display_name: string }>,
    description: '',
    imweb_product_no: '',
    barcode: '',
    supplier_id: '',
    // EVENT 전용 (category='EVENT'일 때만 tags로 저장 — 고객 접수폼 그룹핑용)
    event_type: 'blunt',  // blunt | thinning | long | dry
    spec: '',             // 인치(6.0) / 감모(21to29) / DRY서브(stroke)
  });

  async function handleSubmit() {
    if (!form.sku.trim() || !form.name.trim()) return;

    // price_groups JSONB 조립 (0 이상인 그룹만)
    const pgPayload: Record<string, { price?: number; display_name?: string }> = {};
    for (const [key, val] of Object.entries(form.price_group_values)) {
      if (val.price > 0 || val.display_name) {
        pgPayload[key] = { price: val.price || undefined, display_name: val.display_name || undefined };
      }
    }

    await createProduct.mutateAsync({
      sku: form.sku.trim(),
      name: form.name.trim(),
      category: form.category,
      price: form.price,
      price_dealer: form.price_group_values['dealer']?.price || undefined,  // dual-write
      price_academy: form.price_group_values['academy']?.price || undefined, // dual-write
      price_purchase: form.price_purchase || undefined,
      price_groups: pgPayload,
      description: form.description.trim() || undefined,
      imweb_product_no: form.imweb_product_no.trim() || undefined,
      barcode: form.barcode.trim() || undefined,
      supplier_id: form.supplier_id || undefined,
      ...(form.category === 'EVENT'
        ? { tags: { event_type: form.event_type, spec: form.spec.trim() } }
        : {}),
    });

    router.push('/products');
  }

  async function checkSkuDuplicate(sku: string) {
    if (!sku.trim()) { setSkuStatus('idle'); return; }
    setSkuStatus('checking');
    try {
      const res = await fetch(`/api/products?sku_check=${encodeURIComponent(sku.trim())}`);
      const data = await res.json();
      setSkuStatus(data.exists ? 'duplicate' : 'available');
    } catch {
      setSkuStatus('idle');
    }
  }

  return (
    <>
      <Topbar title="제품 등록" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <Button variant="ghost" size="sm" onClick={() => router.push('/products')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>

        <Card>
          <h3 className="text-sm font-bold text-stone-900 mb-4">기본 정보</h3>
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">SKU *</label>
                <input
                  type="text"
                  value={form.sku}
                  onChange={(e) => { setForm({ ...form, sku: e.target.value }); setSkuStatus('idle'); }}
                  onBlur={() => checkSkuDuplicate(form.sku)}
                  placeholder="MAM-BL-060"
                  className={`w-full h-9 px-3 rounded-lg border text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400 ${
                    skuStatus === 'duplicate' ? 'border-red-400 bg-red-50' : 'border-neutral-200 bg-stone-50'
                  }`}
                />
                {skuStatus === 'checking' && <p className="text-xs text-neutral-400 mt-0.5">확인 중...</p>}
                {skuStatus === 'available' && (
                  <p className="text-xs text-green-600 mt-0.5 flex items-center gap-1"><CheckCircle size={11} />사용 가능</p>
                )}
                {skuStatus === 'duplicate' && (
                  <p className="text-xs text-red-500 mt-0.5 flex items-center gap-1"><XCircle size={11} />이미 사용 중인 SKU입니다</p>
                )}
              </div>
              <div>
                <label className="text-xs text-neutral-500">카테고리</label>
                <select
                  value={form.category}
                  onChange={(e) => setForm({ ...form, category: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400"
                >
                  {CATEGORY_OPTIONS.map((c) => (
                    <option key={c.value} value={c.value}>{c.label}</option>
                  ))}
                </select>
              </div>
            </div>

            <div>
              <label className="text-xs text-neutral-500">제품명 *</label>
              <input
                type="text"
                value={form.name}
                onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="마모루 블런트 6.0"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
              />
            </div>

            {/* EVENT 전용 — 고객 접수폼 그룹핑용 (종류·옵션) */}
            {form.category === 'EVENT' && (
              <div className="rounded-lg border border-indigo-100 bg-indigo-50/50 p-3 space-y-3">
                <p className="text-[11px] text-indigo-700">고객 접수폼에서 이 값으로 그룹핑됩니다. (재고는 아래 가격/재고 입력, 마감임박은 재고 ≤3일 때 자동 표시)</p>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-xs text-neutral-500">종류</label>
                    <select
                      value={form.event_type}
                      onChange={(e) => setForm({ ...form, event_type: e.target.value })}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-stone-400"
                    >
                      <option value="blunt">블런트</option>
                      <option value="thinning">틴닝</option>
                      <option value="long">장가위</option>
                      <option value="dry">DRY</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-xs text-neutral-500">옵션(spec)</label>
                    <input
                      type="text"
                      value={form.spec}
                      onChange={(e) => setForm({ ...form, spec: e.target.value })}
                      placeholder={form.event_type === 'thinning' ? 'under20 / 21to29 / over30' : form.event_type === 'dry' ? 'stroke / slicing / curve' : '6.0 (인치)'}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
                    />
                  </div>
                </div>
              </div>
            )}
          </div>
        </Card>

        <Card>
          <h3 className="text-sm font-bold text-stone-900 mb-4">가격 정보</h3>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div>
              <label className="text-xs text-neutral-500">소매가</label>
              <input
                type="number"
                value={form.price || ''}
                onChange={(e) => setForm({ ...form, price: parseInt(e.target.value) || 0 })}
                placeholder="0"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
              />
            </div>
            {/* 동적 단가 그룹 */}
            {Object.entries(priceGroups).map(([key, def]) => (
              <div key={key}>
                <label className="text-xs text-neutral-500">{def.label}</label>
                <input
                  type="number"
                  value={form.price_group_values[key]?.price || ''}
                  onChange={(e) => setForm({ ...form, price_group_values: { ...form.price_group_values, [key]: { ...form.price_group_values[key], price: parseInt(e.target.value) || 0, display_name: form.price_group_values[key]?.display_name || '' } } })}
                  placeholder="0"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
                />
              </div>
            ))}
            <div>
              <label className="text-xs text-neutral-500">매입가</label>
              <input
                type="number"
                value={form.price_purchase || ''}
                onChange={(e) => setForm({ ...form, price_purchase: parseInt(e.target.value) || 0 })}
                placeholder="0"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
              />
            </div>
          </div>
          {/* 단가 그룹별 납품명 */}
          {Object.entries(priceGroups).length > 0 && (
            <div className="mt-3 grid grid-cols-2 gap-3">
              {Object.entries(priceGroups).map(([key, def]) => (
                <div key={`dn-${key}`}>
                  <label className="text-xs text-neutral-500">{def.label} 납품명</label>
                  <input
                    type="text"
                    value={form.price_group_values[key]?.display_name || ''}
                    onChange={(e) => setForm({ ...form, price_group_values: { ...form.price_group_values, [key]: { ...form.price_group_values[key], price: form.price_group_values[key]?.price || 0, display_name: e.target.value } } })}
                    placeholder="미입력 시 기본 제품명 사용"
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
                  />
                </div>
              ))}
            </div>
          )}
        </Card>

        <Card>
          <h3 className="text-sm font-bold text-stone-900 mb-4">추가 정보</h3>
          <div className="space-y-3">
            <div>
              <label className="text-xs text-neutral-500">매입처</label>
              <SupplierSelect
                value={form.supplier_id}
                onChange={(id) => setForm({ ...form, supplier_id: id })}
                placeholder="매입처 선택 (선택)"
              />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">아임웹 상품번호</label>
                <input
                  type="text"
                  value={form.imweb_product_no}
                  onChange={(e) => setForm({ ...form, imweb_product_no: e.target.value })}
                  placeholder="아임웹 상품관리에서 확인"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
                />
              </div>
              <div>
                <label className="text-xs text-neutral-500">바코드</label>
                <input
                  type="text"
                  value={form.barcode}
                  onChange={(e) => setForm({ ...form, barcode: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400"
                />
              </div>
            </div>

            <div>
              <label className="text-xs text-neutral-500">설명</label>
              <textarea
                value={form.description}
                onChange={(e) => setForm({ ...form, description: e.target.value })}
                rows={3}
                placeholder="제품 설명 (선택)"
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400 resize-none"
              />
            </div>
          </div>
        </Card>

        <Button
          className="w-full"
          disabled={!form.sku.trim() || !form.name.trim() || createProduct.isPending || skuStatus === 'duplicate'}
          onClick={handleSubmit}
        >
          {createProduct.isPending ? '등록 중...' : '제품 등록'}
        </Button>
      </div>
    </>
  );
}
