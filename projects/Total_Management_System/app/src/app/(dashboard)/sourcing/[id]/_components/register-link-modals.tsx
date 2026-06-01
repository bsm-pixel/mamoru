'use client';

import { useState, useMemo } from 'react';
import toast from 'react-hot-toast';
import { X, Loader2, Search } from 'lucide-react';
import { useCreateProduct } from '@/hooks/use-product-detail';
import { useLinkSourcingItem, type SourcingItem } from '@/hooks/use-sourcing';
import { useProducts } from '@/hooks/use-sales';
import { useSetting } from '@/hooks/use-settings';

type Mode = 'product' | 'supply' | 'link';

const inp =
  'w-full px-3 py-2 text-sm rounded-lg border border-neutral-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-black/10 focus:border-neutral-400';

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="block">
      <span className="block text-[11px] text-neutral-500 mb-1 font-medium">{label}</span>
      {children}
    </label>
  );
}

/** 소싱 품목 → 제품/부자재 등록 또는 기존 연결 모달 */
export function RegisterLinkModal({ mode, item, poId, exchangeRate, onClose }: {
  mode: Mode;
  item: SourcingItem;
  poId: string;
  exchangeRate: number;
  onClose: () => void;
}) {
  const link = useLinkSourcingItem(poId);
  const krw = Math.round((item.unit_price || 0) * (exchangeRate || 0));
  // mock: 프리픽스(데모 그라데이션)는 실제 URL 아니므로 제외
  const firstPhoto = (item.inbound_photos ?? []).find((p) => p && !p.startsWith('mock:')) || null;

  return (
    <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 bg-white border-b border-neutral-200 px-5 py-3 flex items-center justify-between">
          <span className="font-bold text-indigo-black">
            {mode === 'product' ? '제품으로 등록' : mode === 'supply' ? '부자재로 등록' : '기존 제품 연결'}
          </span>
          <button type="button" onClick={onClose} className="text-neutral-400 hover:text-neutral-700"><X size={18} /></button>
        </div>
        <div className="p-5">
          {mode === 'link' ? (
            <LinkMode item={item} link={link} onClose={onClose} />
          ) : (
            <RegisterMode mode={mode} item={item} krw={krw} firstPhoto={firstPhoto} link={link} onClose={onClose} />
          )}
        </div>
      </div>
    </div>
  );
}

function RegisterMode({ mode, item, krw, firstPhoto, link, onClose }: {
  mode: 'product' | 'supply';
  item: SourcingItem;
  krw: number;
  firstPhoto: string | null;
  link: ReturnType<typeof useLinkSourcingItem>;
  onClose: () => void;
}) {
  const createProduct = useCreateProduct();
  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', { BL: '블런트', TH: '틴닝', LO: '장가위', SL: '드라이' });
  const categories = useSetting<string[]>('inventory.categories', ['BL', 'TH', 'LO', 'SL']);
  const [name, setName] = useState(item.product_name || '');
  const [category, setCategory] = useState(categories[0] || 'BL');
  const [price, setPrice] = useState(krw);
  const [busy, setBusy] = useState(false);

  const submit = async () => {
    if (!name.trim()) { toast.error('이름을 입력하세요'); return; }
    setBusy(true);
    try {
      let productId: string;
      if (mode === 'product') {
        const skuRes = await fetch(`/api/products/next-sku?category=${category}`);
        const skuData = await skuRes.json();
        if (!skuData.sku) throw new Error('SKU 채번 실패');
        const res = await createProduct.mutateAsync({
          sku: skuData.sku,
          name: name.trim(),
          category,
          price: 0,
          price_purchase: price,
          image_url: firstPhoto ?? undefined,
          purchase_url: item.vendor_url ?? undefined,
        });
        productId = res.product.id;
      } else {
        const res = await fetch('/api/supplies', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            name: name.trim(),
            purchase_url: item.vendor_url || undefined,
            price_purchase: price,
            image_url: firstPhoto || undefined,
            memo: item.features_memo || undefined,
          }),
        });
        if (!res.ok) throw new Error(await res.text());
        const data = await res.json();
        productId = data.supply.id;
        toast.success('부자재 등록 완료');
      }
      await link.mutateAsync({ itemId: item.id, linked_product_id: productId });
      onClose();
    } catch (e) {
      toast.error('등록 실패: ' + (e instanceof Error ? e.message : String(e)));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="space-y-3">
      {firstPhoto && (
        // eslint-disable-next-line @next/next/no-img-element
        <img src={firstPhoto} alt="" className="w-20 h-20 rounded-lg object-cover border border-neutral-200" />
      )}
      <Field label="이름"><input value={name} onChange={(e) => setName(e.target.value)} className={inp} /></Field>
      {mode === 'product' && (
        <Field label="카테고리">
          <select value={category} onChange={(e) => setCategory(e.target.value)} className={inp}>
            {categories.map((c) => <option key={c} value={c}>{catLabels[c] || c}</option>)}
          </select>
        </Field>
      )}
      <Field label={`매입가 (₩) · 단가 ¥${item.unit_price} 환산`}>
        <input type="number" value={price || ''} onChange={(e) => setPrice(Number(e.target.value) || 0)} className={inp} />
      </Field>
      <div className="text-[11px] text-neutral-500 leading-relaxed">
        업체명·회사링크·품목링크는 <strong>소싱에 보존</strong>되어 {mode === 'product' ? '제품' : '부자재'} 상세에서 표시됩니다. (제품 원본은 안 바뀜)
      </div>
      <button
        type="button"
        onClick={submit}
        disabled={busy}
        className="w-full py-2.5 rounded-lg bg-indigo-black text-white text-sm font-bold disabled:opacity-50 inline-flex items-center justify-center gap-2 hover:opacity-90"
      >
        {busy && <Loader2 size={15} className="animate-spin" />}
        {mode === 'product' ? '제품 등록 + 연결' : '부자재 등록 + 연결'}
      </button>
    </div>
  );
}

function LinkMode({ item, link, onClose }: {
  item: SourcingItem;
  link: ReturnType<typeof useLinkSourcingItem>;
  onClose: () => void;
}) {
  const { data: products = [] } = useProducts({ includeInactive: false });
  const [q, setQ] = useState(item.product_name || '');
  const [busy, setBusy] = useState(false);

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    if (!s) return products.slice(0, 40);
    return products
      .filter((p) => (p.name || '').toLowerCase().includes(s) || (p.sku || '').toLowerCase().includes(s))
      .slice(0, 40);
  }, [products, q]);

  const pick = async (productId: string) => {
    setBusy(true);
    try {
      await link.mutateAsync({ itemId: item.id, linked_product_id: productId });
      toast.success('연결 완료');
      onClose();
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="space-y-3">
      <div className="relative">
        <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="제품명·SKU 검색" className={inp + ' pl-8'} autoFocus />
      </div>
      <div className="max-h-72 overflow-y-auto divide-y divide-neutral-100 border border-neutral-200 rounded-lg">
        {filtered.length === 0 ? (
          <div className="p-4 text-center text-xs text-neutral-400">검색 결과 없음</div>
        ) : (
          filtered.map((p) => (
            <button
              key={p.id}
              type="button"
              onClick={() => pick(p.id)}
              disabled={busy}
              className="w-full flex items-center gap-2 p-2.5 hover:bg-neutral-50 text-left disabled:opacity-50"
            >
              <span className="font-mono text-[10px] text-neutral-400 w-24 truncate">{p.sku}</span>
              <span className="flex-1 text-sm text-indigo-black truncate">{p.name}</span>
              <span className="text-[10px] text-neutral-400 flex-shrink-0">{p.category}</span>
            </button>
          ))
        )}
      </div>
    </div>
  );
}
