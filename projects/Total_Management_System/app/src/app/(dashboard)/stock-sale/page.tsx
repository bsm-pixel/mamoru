'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useEvents, useEventPatch, useEventDelete } from '@/hooks/use-events';
import { EVENT_STATUS_LABEL, type EventSubmission, type EventStatus } from '@/lib/event/types';
import { Tag, Loader2, Truck, ExternalLink, X, Package } from 'lucide-react';

/**
 * 재고판매(LS) 어드민 — 2026-07-21 (단계3)
 * event_submissions(kind='stock_sale') 재사용. EVENT 와 같은 파이프라인이라 훅도 공용.
 * 캠페인/할인 없는 단순 목록 — 접수 → 입금확인 → 판매전환.
 */

const TABS: { key: EventStatus; label: string }[] = [
  { key: 'received', label: '신규접수' },
  { key: 'payment_noticed', label: '입금대기' },
  { key: 'converted', label: '판매전환' },
  { key: 'cancelled', label: '취소' },
];

const FORM_URL = 'https://page.mamoru.kr/projects/stock_sale/page_form.html';
const won = (n: number) => `${(n || 0).toLocaleString()}원`;
const fmtPhone = (p: string | null) => (p || '').replace(/(\d{3})(\d{3,4})(\d{4})/, '$1-$2-$3');

function Row({ label, value }: { label: string; value: React.ReactNode }) {
  return (
    <div className="flex justify-between gap-2 py-1.5 border-b border-neutral-50 text-sm">
      <span className="text-neutral-400 shrink-0 whitespace-nowrap">{label}</span>
      <span className="text-right text-neutral-800 min-w-0">{value}</span>
    </div>
  );
}

export default function StockSalePage() {
  const router = useRouter();
  const [tab, setTab] = useState<EventStatus>('received');
  const [selId, setSelId] = useState<string | null>(null);
  const { data: all, isLoading } = useEvents('all', 'stock_sale');
  const patch = useEventPatch();
  const del = useEventDelete();

  const counts = useMemo(() => {
    const m: Record<string, number> = {};
    (all || []).forEach((e) => { m[e.status] = (m[e.status] || 0) + 1; });
    return m;
  }, [all]);
  const list = useMemo(() => (all || []).filter((e) => e.status === tab), [all, tab]);
  const sel = useMemo(() => (all || []).find((e) => e.id === selId) || null, [all, selId]);

  return (
    <>
      <Topbar title="재고판매" />
      <div className="min-h-screen bg-neutral-50 px-4 md:px-6 py-4 space-y-4 overflow-x-hidden">
        <div className="flex items-start gap-2 rounded-xl bg-amber-50 border border-amber-100 px-3 py-2.5 text-xs text-amber-800">
          <Tag size={14} className="shrink-0 mt-0.5" />
          <p>
            사무실 재고를 아임웹 카탈로그로 판매합니다. 제품은 <button onClick={() => router.push('/products')} className="underline font-semibold">제품관리</button>에서
            카테고리 <b>재고판매</b>로 등록하세요. 입금확인 시 자동으로 <button onClick={() => router.push('/sales')} className="underline font-semibold">판매관리</button>로 넘어갑니다.
          </p>
        </div>

        <div className="flex items-center justify-between gap-2 flex-wrap">
          <h2 className="text-lg font-bold text-neutral-900">주문 접수</h2>
          <a href={FORM_URL} target="_blank" rel="noreferrer"
            className="inline-flex items-center gap-1 text-[11px] font-semibold text-amber-700 bg-amber-50 hover:bg-amber-100 px-2.5 py-1 rounded-full transition">
            <ExternalLink size={11} />고객 주문 페이지 열기
          </a>
        </div>

        {/* 탭 */}
        <div className="flex gap-1 overflow-x-auto scrollbar-hide border-b border-neutral-200">
          {TABS.map((t) => (
            <button key={t.key} onClick={() => setTab(t.key)}
              className={`shrink-0 whitespace-nowrap px-3 py-2.5 text-sm font-semibold border-b-2 transition ${
                tab === t.key ? 'border-neutral-900 text-neutral-900' : 'border-transparent text-neutral-400 hover:text-neutral-600'
              }`}>
              {t.label}
              {(counts[t.key] || 0) > 0 && <span className="ml-1.5 text-xs px-1.5 py-0.5 rounded-full bg-neutral-200 text-neutral-600">{counts[t.key]}</span>}
            </button>
          ))}
        </div>

        {isLoading ? (
          <div className="py-16 text-center text-sm text-neutral-400"><Loader2 size={20} className="animate-spin inline" /></div>
        ) : list.length === 0 ? (
          <div className="py-16 text-center text-sm text-neutral-400">{EVENT_STATUS_LABEL[tab]} 건이 없습니다</div>
        ) : (
          <div className="space-y-2">
            {list.map((e) => (
              <button key={e.id} onClick={() => setSelId(e.id)}
                className="w-full text-left bg-white rounded-xl border border-neutral-200 px-4 py-3 hover:border-neutral-400 transition">
                <div className="flex items-center justify-between gap-2">
                  <div className="min-w-0">
                    <div className="text-[10px] font-mono text-neutral-400">{e.event_number}</div>
                    <div className="text-sm font-bold text-neutral-900 truncate">{e.customer_name}</div>
                  </div>
                  <div className="shrink-0 text-right">
                    <div className="text-sm font-bold text-neutral-900">{won(e.total_amount)}</div>
                    <div className="text-[11px] text-neutral-400 flex items-center gap-0.5 justify-end"><Truck size={11} />택배</div>
                  </div>
                </div>
                <div className="mt-1.5 text-xs text-neutral-500 truncate">
                  {(e.items || []).map((it) => `${it.product_name}×${it.qty}`).join(', ')}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>

      <SlidePanel open={!!sel} onClose={() => setSelId(null)} title="재고판매 접수 상세" className="sm:w-[440px]">
        {sel && <StockDetail ev={sel} patch={patch} del={del} onDone={() => setSelId(null)} goSales={() => router.push('/sales')} />}
      </SlidePanel>
    </>
  );
}

function StockDetail({ ev, patch, del, onDone, goSales }: {
  ev: EventSubmission;
  patch: ReturnType<typeof useEventPatch>;
  del: ReturnType<typeof useEventDelete>;
  onDone: () => void;
  goSales: () => void;
}) {
  const confirmPay = () => {
    if (!window.confirm(`${ev.customer_name}님 입금을 확인하고 판매로 전환합니다. (재고가 차감됩니다)`)) return;
    patch.mutate({ id: ev.id, action: 'confirm_payment' }, { onSuccess: onDone });
  };

  const itemsSum = (ev.items || []).reduce((s, it) => s + it.unit_price * it.qty, 0);
  const shipFee = ev.total_amount - itemsSum;   // 배송비(상품 외 가산액)

  return (
    <div className="space-y-4">
      <div>
        <div className="text-[11px] font-mono text-neutral-400">{ev.event_number}</div>
        <div className="text-lg font-bold text-neutral-900">{ev.customer_name}</div>
        <div className="text-sm text-neutral-500">{fmtPhone(ev.customer_phone)}</div>
        <span className="inline-block mt-1 text-[11px] px-2 py-0.5 rounded-full bg-neutral-100 text-neutral-600">{EVENT_STATUS_LABEL[ev.status]}</span>
      </div>

      <div className="rounded-xl bg-neutral-50 p-3 space-y-1">
        {(ev.items || []).map((it, i) => (
          <div key={i} className="flex justify-between text-sm">
            <span className="text-neutral-700 min-w-0 truncate">{it.product_name} ×{it.qty}</span>
            <span className="shrink-0 text-neutral-900">{won(it.unit_price * it.qty)}</span>
          </div>
        ))}
        {shipFee > 0 && (
          <div className="flex justify-between text-sm text-neutral-500">
            <span>배송비</span><span>{won(shipFee)}</span>
          </div>
        )}
        <div className="flex justify-between pt-1 mt-1 border-t border-neutral-200 text-sm font-bold">
          <span>합계</span><span>{won(ev.total_amount)}</span>
        </div>
      </div>

      <div>
        <Row label="수령방법" value="택배 발송" />
        <Row label="배송지" value={[ev.postcode ? `(${ev.postcode})` : '', ev.address1, ev.address2].filter(Boolean).join(' ') || '-'} />
        {ev.memo && <Row label="메모" value={ev.memo} />}
      </div>

      {/* 액션 */}
      <div className="space-y-2 pt-2">
        {(ev.status === 'received' || ev.status === 'payment_noticed') && (
          <>
            <button
              disabled={patch.isPending}
              onClick={confirmPay}
              className="w-full py-2.5 rounded-lg bg-emerald-600 text-white text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
            >
              {patch.isPending ? <Loader2 size={16} className="animate-spin" /> : '입금확인 → 판매 전환'}
            </button>
            {ev.status === 'received' && (
              <button
                disabled={patch.isPending}
                onClick={() => patch.mutate({ id: ev.id, action: 'payment_notice' }, { onSuccess: onDone })}
                className="w-full py-2 rounded-lg border border-neutral-300 text-neutral-600 text-xs font-semibold disabled:opacity-50 flex items-center justify-center gap-2"
              >
                <Package size={14} /> 입금안내 재발송 (선택)
              </button>
            )}
          </>
        )}
        {ev.status === 'converted' && (
          <button onClick={goSales} className="w-full py-2.5 rounded-lg border border-neutral-300 text-sm font-semibold text-neutral-700">
            판매로 전환됨 — 판매관리에서 송장·발송 처리 →
          </button>
        )}

        {ev.status !== 'cancelled' && (
          <button
            disabled={patch.isPending}
            onClick={() => {
              const msg = ev.status === 'converted'
                ? '이 접수를 취소 처리합니다.\n※ 판매 건은 판매관리에서 별도로 취소하세요 — 여기선 접수 기록만 취소됩니다.'
                : '이 접수를 취소합니다.';
              if (!window.confirm(msg)) return;
              patch.mutate({ id: ev.id, action: 'cancel' }, { onSuccess: onDone });
            }}
            className="w-full py-2 rounded-lg text-xs text-neutral-400 hover:text-red-500"
          >
            접수 취소
          </button>
        )}

        {(ev.status === 'cancelled' || ev.status === 'converted') && (
          <button
            disabled={del.isPending}
            onClick={() => {
              if (!window.confirm('이 접수 기록을 완전히 삭제합니다. 되돌릴 수 없습니다.\n※ 연결된 판매 건은 판매관리에서 별도 관리됩니다.')) return;
              del.mutate({ id: ev.id }, { onSuccess: onDone });
            }}
            className="w-full py-2 rounded-lg text-xs font-semibold text-red-500 border border-red-200 hover:bg-red-50 disabled:opacity-50 flex items-center justify-center gap-1.5"
          >
            {del.isPending ? <Loader2 size={13} className="animate-spin" /> : <X size={13} />}
            접수 기록 삭제
          </button>
        )}
      </div>
    </div>
  );
}
