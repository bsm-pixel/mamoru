'use client';

import { use, useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useProduct, useUpdateProduct } from '@/hooks/use-product-detail';
import { formatKRW } from '@/lib/utils/format';
import { SupplierSelect } from '@/components/ui/supplier-select';
import { ArrowLeft, Save, Package, Hash } from 'lucide-react';

const CATEGORY_LABEL: Record<string, string> = {
  BL: '블런트',
  TH: '틴닝',
  LO: '장가위',
  SL: '슬라이싱',
};

const CATEGORY_OPTIONS = [
  { value: 'BL', label: '블런트' },
  { value: 'TH', label: '틴닝' },
  { value: 'LO', label: '장가위' },
  { value: 'SL', label: '슬라이싱' },
];

export default function ProductDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();
  const { data, isLoading } = useProduct(id);
  const updateProduct = useUpdateProduct();

  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState({
    name: '',
    category: '',
    price: 0,
    price_dealer: 0,
    price_purchase: 0,
    description: '',
    imweb_product_no: '',
    barcode: '',
    supplier_id: '',
  });

  useEffect(() => {
    if (data?.product && !editing) {
      const p = data.product;
      setForm({
        name: p.name,
        category: p.category,
        price: p.price,
        price_dealer: p.price_dealer || 0,
        price_purchase: p.price_purchase || 0,
        description: p.description || '',
        imweb_product_no: p.imweb_product_no || '',
        barcode: p.barcode || '',
        supplier_id: p.supplier_id || '',
      });
    }
  }, [data, editing]);

  async function handleSave() {
    await updateProduct.mutateAsync({
      id,
      name: form.name,
      category: form.category,
      price: form.price,
      price_dealer: form.price_dealer,
      price_purchase: form.price_purchase,
      description: form.description || null,
      imweb_product_no: form.imweb_product_no || null,
      barcode: form.barcode || null,
      supplier_id: form.supplier_id || null,
    });
    setEditing(false);
  }

  if (isLoading) {
    return (
      <>
        <Topbar title="제품 상세" />
        <div className="px-4 md:px-6 py-4 space-y-4">
          <Skeleton className="h-48" />
          <Skeleton className="h-32" />
        </div>
      </>
    );
  }

  if (!data?.product) {
    return (
      <>
        <Topbar title="제품 상세" />
        <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
          제품 정보를 찾을 수 없습니다
        </div>
      </>
    );
  }

  const { product: p, supplier } = data;

  return (
    <>
      <Topbar title="제품 상세" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="flex items-center gap-2">
          <Button variant="ghost" size="sm" onClick={() => router.push('/products')}>
            <ArrowLeft size={14} />
            목록
          </Button>
          <Button variant="ghost" size="sm" onClick={() => router.push(`/products/${id}/serials`)}>
            <Hash size={14} />
            시리얼 관리
          </Button>
        </div>

        {/* 제품 정보 */}
        <Card>
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <Package size={18} className="text-terracotta" />
              <h3 className="text-sm font-bold text-indigo-black">{p.name}</h3>
              <Badge className="bg-neutral-100 text-neutral-600">{p.sku}</Badge>
            </div>
            {!editing ? (
              <Button variant="ghost" size="sm" onClick={() => setEditing(true)}>수정</Button>
            ) : (
              <div className="flex gap-1">
                <Button variant="ghost" size="sm" onClick={() => setEditing(false)}>취소</Button>
                <Button size="sm" onClick={handleSave} disabled={updateProduct.isPending}>
                  <Save size={14} />
                  {updateProduct.isPending ? '저장 중...' : '저장'}
                </Button>
              </div>
            )}
          </div>

          {editing ? (
            <div className="space-y-3">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-neutral-500">제품명</label>
                  <input
                    type="text"
                    value={form.name}
                    onChange={(e) => setForm({ ...form, name: e.target.value })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
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

              {/* 3단 가격 */}
              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="text-xs text-neutral-500">소매가</label>
                  <input
                    type="number"
                    value={form.price || ''}
                    onChange={(e) => setForm({ ...form, price: parseInt(e.target.value) || 0 })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">도매가</label>
                  <input
                    type="number"
                    value={form.price_dealer || ''}
                    onChange={(e) => setForm({ ...form, price_dealer: parseInt(e.target.value) || 0 })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">매입가</label>
                  <input
                    type="number"
                    value={form.price_purchase || ''}
                    onChange={(e) => setForm({ ...form, price_purchase: parseInt(e.target.value) || 0 })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-neutral-500">아임웹 상품번호</label>
                  <input
                    type="text"
                    value={form.imweb_product_no}
                    onChange={(e) => setForm({ ...form, imweb_product_no: e.target.value })}
                    placeholder="아임웹 상품 URL의 번호"
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">바코드</label>
                  <input
                    type="text"
                    value={form.barcode}
                    onChange={(e) => setForm({ ...form, barcode: e.target.value })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
              </div>

              <div>
                <label className="text-xs text-neutral-500">매입처</label>
                <SupplierSelect
                  value={form.supplier_id}
                  displayName={supplier?.name}
                  onChange={(id) => setForm({ ...form, supplier_id: id })}
                  placeholder="매입처 선택"
                />
              </div>

              <div>
                <label className="text-xs text-neutral-500">설명</label>
                <textarea
                  value={form.description}
                  onChange={(e) => setForm({ ...form, description: e.target.value })}
                  rows={3}
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none"
                />
              </div>
            </div>
          ) : (
            <>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 text-sm">
                <div>
                  <span className="text-xs text-neutral-500">카테고리</span>
                  <p>{CATEGORY_LABEL[p.category] || p.category}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">재고</span>
                  <p className={p.stock_quantity > 0 ? 'font-bold' : 'font-bold text-red-500'}>
                    {p.stock_quantity}개
                  </p>
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

              {/* 3단 가격 */}
              <div className="mt-4 pt-3 border-t border-neutral-100">
                <h4 className="text-xs text-neutral-500 mb-2">가격 정보</h4>
                <div className="grid grid-cols-3 gap-3">
                  <div>
                    <p className="text-xs text-neutral-400">소매가</p>
                    <p className="text-sm font-bold text-terracotta">{formatKRW(p.price)}</p>
                  </div>
                  <div>
                    <p className="text-xs text-neutral-400">도매가</p>
                    <p className="text-sm font-bold text-purple-600">{p.price_dealer > 0 ? formatKRW(p.price_dealer) : '-'}</p>
                  </div>
                  <div>
                    <p className="text-xs text-neutral-400">매입가</p>
                    <p className="text-sm font-bold text-neutral-600">{p.price_purchase > 0 ? formatKRW(p.price_purchase) : '-'}</p>
                  </div>
                </div>
              </div>

              {/* 매입처 */}
              {supplier && (
                <div className="mt-3 pt-3 border-t border-neutral-100">
                  <span className="text-xs text-neutral-500">매입처</span>
                  <p className="text-sm">{supplier.name}</p>
                </div>
              )}

              {/* 아임웹 매핑 */}
              {p.imweb_product_no && (
                <div className="mt-3 pt-3 border-t border-neutral-100">
                  <span className="text-xs text-neutral-500">아임웹 상품번호</span>
                  <p className="text-sm font-mono">#{p.imweb_product_no}</p>
                </div>
              )}

              {/* 설명 */}
              {p.description && (
                <div className="mt-3 pt-3 border-t border-neutral-100">
                  <span className="text-xs text-neutral-500">설명</span>
                  <p className="text-sm text-neutral-600 whitespace-pre-wrap">{p.description}</p>
                </div>
              )}
            </>
          )}
        </Card>
      </div>
    </>
  );
}
