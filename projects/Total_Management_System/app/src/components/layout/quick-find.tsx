'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { useRouter } from 'next/navigation';
import { useQuery } from '@tanstack/react-query';
import { Search, User, Receipt, Package, LayoutGrid, CornerDownLeft } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';
import { NAV_GROUPS } from '@/lib/utils/constants';
import { matchSearch } from '@/lib/utils/search-match';
import { formatPhone, formatDate } from '@/lib/utils/format';
import { useProducts } from '@/hooks/use-sales';
import { askBeforeLeave } from '@/hooks/use-unsaved-guard';
import { backdropClose } from '@/lib/ui/backdrop';
import { EscClose } from '@/components/ui/esc-close';

/**
 * 빠른 찾기 (2026-11-15) — Ctrl+K 또는 머리줄의 [찾기]
 *
 * 고객 이름 · 전화 뒤 4자리 · 판매번호 · 제품 · 메뉴 이름을 한 칸에서 찾아 Enter 로 바로 연다.
 * 대시보드 레이아웃에 한 번만 둔다. 여는 쪽은 openQuickFind() 를 부른다.
 */
const OPEN_EVENT = 'mm:quick-find';
export function openQuickFind() {
  if (typeof window !== 'undefined') window.dispatchEvent(new CustomEvent(OPEN_EVENT));
}

/** 제품은 상세 주소가 따로 없어, 제품 화면이 열리면서 이 값을 읽어 그 제품을 편다 */
export const OPEN_PRODUCT_KEY = 'mm:open-product';

interface Hit { key: string; kind: 'menu' | 'customer' | 'sale' | 'product'; title: string; sub?: string; href: string; productId?: string }

const KIND_LABEL: Record<Hit['kind'], string> = { menu: '메뉴', customer: '고객', sale: '판매', product: '제품' };
const KIND_ICON = { menu: LayoutGrid, customer: User, sale: Receipt, product: Package };

function useDebounced(v: string, ms: number) {
  const [d, setD] = useState(v);
  useEffect(() => { const t = setTimeout(() => setD(v), ms); return () => clearTimeout(t); }, [v, ms]);
  return d;
}

export function QuickFind() {
  const router = useRouter();
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');
  const [active, setActive] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const activeRef = useRef<HTMLLIElement>(null);
  const dq = useDebounced(q.trim(), 200);

  useEffect(() => {
    const onOpen = () => { setOpen(true); setQ(''); setActive(0); };
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setOpen((o) => { if (!o) { setQ(''); setActive(0); } return !o; });
      }
    };
    window.addEventListener(OPEN_EVENT, onOpen);
    window.addEventListener('keydown', onKey, true);
    return () => { window.removeEventListener(OPEN_EVENT, onOpen); window.removeEventListener('keydown', onKey, true); };
  }, []);
  useEffect(() => { if (open) setTimeout(() => inputRef.current?.focus(), 0); }, [open]);
  // 열려 있는 동안 초점이 뒤쪽 화면으로 넘어가면 찾기 칸으로 되돌린다(초점 가두기)
  useEffect(() => {
    if (!open) return;
    const keep = (e: FocusEvent) => {
      const t = e.target as Node | null;
      if (t && inputRef.current && !inputRef.current.closest('[role=dialog]')?.contains(t)) inputRef.current.focus();
    };
    document.addEventListener('focusin', keep);
    return () => document.removeEventListener('focusin', keep);
  }, [open]);
  useEffect(() => { activeRef.current?.scrollIntoView({ block: 'nearest' }); }, [active]);

  const { data: products = [] } = useProducts();

  // 고객: 이름 · 전화(뒤 4자리 포함) — 서버 검색
  const { data: customers = [], isFetching: loadingC } = useQuery({
    queryKey: ['quick-find-customers', dq],
    enabled: open && dq.length >= 2,
    staleTime: 30_000,
    queryFn: async () => {
      const res = await fetch(`/api/customers/search?q=${encodeURIComponent(dq)}`);
      if (!res.ok) return [];
      const j = await res.json();
      return (j.customers || []) as Array<{ id: string; name: string; phone: string | null; company_name?: string | null }>;
    },
  });
  // 판매: 판매번호 · 고객 이름 · 전화
  const { data: sales = [], isFetching: loadingS } = useQuery({
    queryKey: ['quick-find-sales', dq],
    enabled: open && dq.length >= 2,
    staleTime: 30_000,
    queryFn: async () => {
      const safe = dq.replace(/[%,()]/g, '');
      if (!safe) return [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data } = await (createClient() as any)
        .from('offline_sales')
        .select('id, sale_number, customer_name, sale_date, cancelled_at')
        .or(`sale_number.ilike.%${safe}%,customer_name.ilike.%${safe}%,customer_phone.ilike.%${safe}%`)
        .order('sale_date', { ascending: false })
        .limit(6);
      return (data || []) as Array<{ id: string; sale_number: string; customer_name: string; sale_date: string; cancelled_at: string | null }>;
    },
  });

  const hits = useMemo<Hit[]>(() => {
    const out: Hit[] = [];
    const menus = NAV_GROUPS.flatMap((g) => g.items);
    for (const m of menus) if (!q.trim() || matchSearch(q, m.label)) out.push({ key: 'm' + m.href, kind: 'menu', title: m.label, href: m.href });
    if (!q.trim()) return out.slice(0, 12);
    const menuHits = out.slice(0, 5);
    const cust = customers.slice(0, 6).map<Hit>((c) => ({ key: 'c' + c.id, kind: 'customer', title: c.company_name ? `${c.company_name} · ${c.name}` : c.name, sub: formatPhone(c.phone), href: `/customers/${c.id}` }));
    const sale = sales.map<Hit>((s) => ({ key: 's' + s.id, kind: 'sale', title: `${s.sale_number} · ${s.customer_name || ''}`, sub: `${formatDate(s.sale_date)}${s.cancelled_at ? ' · 취소됨' : ''}`, href: `/sales/${s.id}` }));
    const prod = products.filter((p) => p.category !== 'SUP' && matchSearch(q, p.name, p.sku)).slice(0, 6)
      .map<Hit>((p) => ({ key: 'p' + p.id, kind: 'product', title: p.name, sub: p.sku || undefined, href: '/products', productId: p.id }));
    return [...cust, ...sale, ...prod, ...menuHits];
  }, [q, customers, sales, products]);

  if (!open) return null;

  const go = (h: Hit) => {
    setOpen(false);
    if (h.productId) { try { sessionStorage.setItem(OPEN_PRODUCT_KEY, h.productId); } catch { /* 저장 못 해도 제품 화면은 연다 */ } }
    if (askBeforeLeave(h.href)) return;
    if (h.productId && window.location.pathname === '/products') window.dispatchEvent(new CustomEvent(OPEN_PRODUCT_KEY));
    router.push(h.href);
  };

  const idx = Math.min(active, Math.max(hits.length - 1, 0));
  const loading = (loadingC || loadingS) && dq.length >= 2;

  return (
    <div className="fixed inset-0 z-[150] flex items-start justify-center bg-black/40 p-4 pt-[12vh]" data-modal-open="" {...backdropClose(() => setOpen(false))}>
      <EscClose onClose={() => setOpen(false)} />
      <div role="dialog" aria-modal="true" aria-label="빠른 찾기" data-nodirty="" className="w-full max-w-[520px] rounded-xl bg-white border border-neutral-200 overflow-hidden" onClick={(e) => e.stopPropagation()}>
        <div className="relative border-b border-neutral-100">
          <Search size={16} className="absolute left-4 top-1/2 -translate-y-1/2 text-neutral-400" />
          <input
            ref={inputRef}
            value={q}
            onChange={(e) => { setQ(e.target.value); setActive(0); }}
            onKeyDown={(e) => {
              if (e.nativeEvent.isComposing) return;
              if (e.key === 'ArrowDown') { e.preventDefault(); setActive(Math.min(idx + 1, hits.length - 1)); }
              else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(Math.max(idx - 1, 0)); }
              else if (e.key === 'Enter') { e.preventDefault(); if (hits[idx]) go(hits[idx]); }
            }}
            placeholder="고객 이름 · 전화 뒤 4자리 · 판매번호 · 제품 · 메뉴"
            className="w-full h-12 pl-11 pr-4 text-sm placeholder:text-neutral-400 focus:outline-none"
          />
        </div>
        <ul role="listbox" className="max-h-[52vh] overflow-y-auto py-1">
          {hits.map((h, i) => {
            const Icon = KIND_ICON[h.kind];
            return (
              <li key={h.key} ref={i === idx ? activeRef : undefined} role="option" aria-selected={i === idx}
                onMouseEnter={() => setActive(i)} onMouseDown={(e) => { e.preventDefault(); go(h); }}
                className={`flex items-center gap-3 px-4 py-2 cursor-pointer ${i === idx ? 'bg-stone-100' : ''}`}>
                <Icon size={15} className="text-neutral-400 shrink-0" />
                <span className="min-w-0 flex-1">
                  <span className="block text-sm font-medium text-stone-900 truncate">{h.title}</span>
                  {h.sub && <span className="block text-[11px] text-neutral-400 truncate">{h.sub}</span>}
                </span>
                <span className="text-[10.5px] font-semibold text-neutral-400 shrink-0">{KIND_LABEL[h.kind]}</span>
                {i === idx && <CornerDownLeft size={12} className="text-neutral-400 shrink-0" />}
              </li>
            );
          })}
          {hits.length === 0 && (
            <li className="px-4 py-6 text-center text-xs text-neutral-400">{loading ? '찾는 중…' : q.trim().length < 2 ? '두 글자 이상 넣어 주세요' : '맞는 것이 없습니다'}</li>
          )}
        </ul>
        <p className="px-4 py-2 border-t border-neutral-100 text-[11px] text-neutral-400 hidden md:block">↑ ↓ 이동 · Enter 열기 · Esc 닫기</p>
      </div>
    </div>
  );
}
