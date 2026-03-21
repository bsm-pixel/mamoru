'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useCreateProduct } from '@/hooks/use-product-detail';
import { SupplierSelect } from '@/components/ui/supplier-select';
import { ArrowLeft } from 'lucide-react';

const CATEGORY_OPTIONS = [
  { value: 'BL', label: '블런트' },
  { value: 'TH', label: '틴닝' },
  { value: 'LO', label: '장가위' },
  { value: 'SL', label: '슬라이싱' },
];

export default function NewProductPage() {
  const router = useRouter();
  const createProduct = useCreateProduct();

  const [form, setForm] = useState({
    sku: '',
    name: '',
    category: 'BL',
    price: 0,
    price_dealer: 0,
    price_academy: 0,
    price_purchase: 0,
    description: '',
    imweb_product_no: '',
    barcode: '',
    supplier_id: '',
  });

  async function handleSubmit() {
    if (!form.sku.trim() || !form.name.trim()) return;

    await createProduct.mutateAsync({
      sku: form.sku.trim(),
      name: form.name.trim(),
      category: form.category,
      price: form.price,
      price_dealer: form.price_dealer || undefined,
      price_academy: form.price_academy || undefined,
      price_purchase: form.price_purchase || undefined,
      description: form.description.trim() || undefined,
      imweb_product_no: form.imweb_product_no.trim() || undefined,
      barcode: form.barcode.trim() || undefined,
      supplier_id: form.supplier_id || undefined,
    });

    router.push('/products');
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
          <h3 className="text-sm font-bold text-indigo-black mb-4">기본 정보</h3>
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">SKU *</label>
                <input
                  type="text"
                  value={form.sku}
                  onChange={(e) => setForm({ ...form, sku: e.target.value })}
                  placeholder="MM-BL-001"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                />
              </div>
              <div>
                <label className="text-xs text-neutral-500">카테고리</label>
                <select
                  value={form.category}
                  onChange={(e) => setForm({ ...form, category: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
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
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
          </div>
        </Card>

        <Card>
          <h3 className="text-sm font-bold text-indigo-black mb-4">가격 정보</h3>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div>
              <label className="text-xs text-neutral-500">소매가</label>
              <input
                type="number"
                value={form.price || ''}
                onChange={(e) => setForm({ ...form, price: parseInt(e.target.value) || 0 })}
                placeholder="0"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
            <div>
              <label className="text-xs text-neutral-500">딜러가</label>
              <input
                type="number"
                value={form.price_dealer || ''}
                onChange={(e) => setForm({ ...form, price_dealer: parseInt(e.target.value) || 0 })}
                placeholder="0"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
            <div>
              <label className="text-xs text-neutral-500">아카데미가</label>
              <input
                type="number"
                value={form.price_academy || ''}
                onChange={(e) => setForm({ ...form, price_academy: parseInt(e.target.value) || 0 })}
                placeholder="0"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
            <div>
              <label className="text-xs text-neutral-500">매입가</label>
              <input
                type="number"
                value={form.price_purchase || ''}
                onChange={(e) => setForm({ ...form, price_purchase: parseInt(e.target.value) || 0 })}
                placeholder="0"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
          </div>
        </Card>

        <Card>
          <h3 className="text-sm font-bold text-indigo-black mb-4">추가 정보</h3>
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
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                />
              </div>
              <div>
                <label className="text-xs text-neutral-500">바코드</label>
                <input
                  type="text"
                  value={form.barcode}
                  onChange={(e) => setForm({ ...form, barcode: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
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
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none"
              />
            </div>
          </div>
        </Card>

        <Button
          className="w-full"
          disabled={!form.sku.trim() || !form.name.trim() || createProduct.isPending}
          onClick={handleSubmit}
        >
          {createProduct.isPending ? '등록 중...' : '제품 등록'}
        </Button>
      </div>
    </>
  );
}
