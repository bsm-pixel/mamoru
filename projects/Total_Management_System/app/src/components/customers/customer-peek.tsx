'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import Link from 'next/link';
import { useQuery } from '@tanstack/react-query';
import { formatDate, formatKRW, CUSTOMER_TYPE_LABEL, CUSTOMER_TYPE_COLOR } from '@/lib/utils/format';
import { activitySuffix } from '@/lib/customer/display';
import type { CustomerPeekData } from '@/app/api/customers/[id]/peek/route';

/**
 * 고객 미리 보기 (2026-11-21) — 이름에 마우스를 올리면(휴대폰은 길게 누르면) 작은 카드
 *
 * 카드: 유형 · 미수금 · 최근 거래 3건 · 마지막 메모 1줄 · [고객 열기]
 *  - 이름을 누르는 동작(줄 선택 · 상세 열기)은 그대로 둔다. 카드는 얹히는 것뿐.
 *  - 길게 눌러 카드가 뜬 직후의 클릭 한 번은 삼킨다(손을 떼는 순간 줄이 열리지 않게).
 *  - 자료는 카드가 뜰 때만 부른다(api/customers/[id]/peek, 1분 동안 다시 안 부름).
 *  - customerId 가 없으면 그냥 글자.
 */
const OPEN_DELAY = 400;   // 지나가다 스친 마우스에는 안 뜨게
const CLOSE_DELAY = 180;  // 이름 → 카드로 마우스를 옮길 틈
const LONG_PRESS = 550;
const CARD_W = 288;

export function CustomerPeek({ customerId, className = '', children }: {
  customerId?: string | null;
  className?: string;
  children: React.ReactNode;
}) {
  const anchorRef = useRef<HTMLSpanElement | null>(null);
  const openTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const closeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const swallowClick = useRef(false);
  const [pos, setPos] = useState<{ left: number; top: number; above: boolean } | null>(null);

  const clearTimers = () => {
    if (openTimer.current) clearTimeout(openTimer.current);
    if (closeTimer.current) clearTimeout(closeTimer.current);
    openTimer.current = null; closeTimer.current = null;
  };
  const show = useCallback(() => {
    const el = anchorRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const vw = window.innerWidth, vh = window.innerHeight;
    const left = Math.max(8, Math.min(r.left, vw - CARD_W - 8));
    const above = r.bottom + 240 > vh && r.top > 240;
    setPos({ left, top: above ? r.top - 6 : r.bottom + 6, above });
  }, []);
  const hide = useCallback(() => { clearTimers(); setPos(null); }, []);

  useEffect(() => {
    if (!pos) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') hide(); };
    const onDown = (e: Event) => {
      const t = e.target as Node | null;
      if (t && (anchorRef.current?.contains(t) || document.getElementById('customer-peek-card')?.contains(t))) return;
      hide();
    };
    window.addEventListener('keydown', onKey);
    window.addEventListener('scroll', hide, true);
    window.addEventListener('pointerdown', onDown, true);
    return () => {
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('scroll', hide, true);
      window.removeEventListener('pointerdown', onDown, true);
    };
  }, [pos, hide]);
  useEffect(() => clearTimers, []);

  if (!customerId) return <span className={className}>{children}</span>;

  return (
    <>
      <span
        ref={anchorRef}
        className={className}
        data-customer-peek
        onMouseEnter={() => { clearTimers(); openTimer.current = setTimeout(show, OPEN_DELAY); }}
        onMouseLeave={() => { clearTimers(); closeTimer.current = setTimeout(hide, CLOSE_DELAY); }}
        onTouchStart={() => { clearTimers(); openTimer.current = setTimeout(() => { swallowClick.current = true; show(); }, LONG_PRESS); }}
        onTouchMove={clearTimers}
        onTouchEnd={() => { if (openTimer.current) clearTimeout(openTimer.current); openTimer.current = null; }}
        onContextMenu={(e) => { if (swallowClick.current) e.preventDefault(); }}
        onClickCapture={(e) => { if (swallowClick.current) { swallowClick.current = false; e.preventDefault(); e.stopPropagation(); } }}
      >
        {children}
      </span>
      {pos && createPortal(
        <div
          id="customer-peek-card"
          role="dialog"
          aria-label="고객 미리 보기"
          style={{ position: 'fixed', left: pos.left, top: pos.top, width: CARD_W, transform: pos.above ? 'translateY(-100%)' : undefined, zIndex: 80 }}
          className="rounded-xl border border-neutral-200 bg-white p-3 text-left shadow-[0_8px_24px_rgba(0,0,0,0.10)]"
          onMouseEnter={clearTimers}
          onMouseLeave={() => { clearTimers(); closeTimer.current = setTimeout(hide, CLOSE_DELAY); }}
          onClick={(e) => e.stopPropagation()}
        >
          <PeekBody customerId={customerId} onNavigate={hide} />
        </div>,
        document.body,
      )}
    </>
  );
}

function PeekBody({ customerId, onNavigate }: { customerId: string; onNavigate: () => void }) {
  const { data, isLoading, isError } = useQuery({
    queryKey: ['customer-peek', customerId],
    staleTime: 60_000,
    queryFn: async (): Promise<CustomerPeekData> => {
      const res = await fetch(`/api/customers/${customerId}/peek`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
  });

  if (isLoading) return <p className="py-3 text-center text-xs text-neutral-400">불러오는 중…</p>;
  if (isError || !data) return <p className="py-3 text-center text-xs text-neutral-400">고객 정보를 불러오지 못했습니다</p>;

  const c = data.customer;
  const suffix = activitySuffix(c.activity_name, c.position);
  return (
    <div className="space-y-2.5">
      <div className="flex items-center gap-1.5 min-w-0">
        <span className="text-sm font-bold text-stone-900 truncate">{c.name}</span>
        {suffix && <span className="text-[11px] text-neutral-400 truncate">{suffix}</span>}
        <span className={`ml-auto shrink-0 px-2 py-0.5 rounded text-[10.5px] font-bold ${CUSTOMER_TYPE_COLOR[c.customer_type] || CUSTOMER_TYPE_COLOR.retail}`}>
          {CUSTOMER_TYPE_LABEL[c.customer_type] || '일반'}
        </span>
      </div>
      {c.company_name && <p className="-mt-1.5 text-[11px] text-neutral-500 truncate">{c.company_name}</p>}

      <div className="flex items-center justify-between text-xs">
        <span className="text-neutral-500">미수금</span>
        {c.outstanding_balance > 0
          ? <span className="font-bold tabular-nums text-red-500">{formatKRW(c.outstanding_balance)}</span>
          : <span className="text-neutral-400">없음</span>}
      </div>

      <div className="border-t border-neutral-100 pt-2">
        <p className="mb-1 text-[11px] font-semibold text-neutral-500">최근 거래</p>
        {data.recent.length === 0 ? (
          <p className="text-xs text-neutral-400">거래 기록이 없습니다</p>
        ) : (
          <ul className="space-y-1">
            {data.recent.map((t, i) => (
              <li key={i} className={`flex items-center gap-1.5 text-xs ${t.cancelled ? 'text-neutral-400 line-through' : 'text-stone-700'}`}>
                <span className="shrink-0 w-12 text-[10.5px] font-semibold text-neutral-500 no-underline">{t.kind}</span>
                <span className="shrink-0 tabular-nums text-neutral-500">{formatDate(t.date)}</span>
                <span className="min-w-0 flex-1 truncate text-right tabular-nums">{t.amount != null && t.amount > 0 ? formatKRW(t.amount) : t.label}</span>
              </li>
            ))}
          </ul>
        )}
      </div>

      {data.note && (
        <div className="border-t border-neutral-100 pt-2">
          <p className="mb-0.5 text-[11px] font-semibold text-neutral-500">마지막 메모 <span className="font-normal text-neutral-400">{formatDate(data.note.created_at)}</span></p>
          <p className="text-xs text-stone-700 truncate" title={data.note.body}>{data.note.body}</p>
        </div>
      )}

      <Link
        href={`/customers/${c.id}`}
        onClick={onNavigate}
        className="block rounded-lg bg-stone-900 py-1.5 text-center text-xs font-semibold text-white hover:bg-stone-800 transition"
      >
        고객 열기
      </Link>
    </div>
  );
}
