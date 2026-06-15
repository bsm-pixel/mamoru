'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useEvents, useEventPatch } from '@/hooks/use-events';
import { EVENT_STATUS_LABEL, type EventSubmission, type EventStatus } from '@/lib/event/types';
import { Zap, Loader2, Package, Truck, Store } from 'lucide-react';

const TABS: { key: EventStatus | 'all'; label: string }[] = [
  { key: 'received', label: '신규접수' },
  { key: 'payment_noticed', label: '입금대기' },
  { key: 'converted', label: '판매전환' },
  { key: 'cancelled', label: '취소' },
];

const won = (n: number) => `${(n || 0).toLocaleString()}원`;
const fmtPhone = (p: string | null) => (p || '').replace(/(\d{3})(\d{3,4})(\d{4})/, '$1-$2-$3');

export default function EventsPage() {
  const router = useRouter();
  const [tab, setTab] = useState<EventStatus>('received');
  const [selId, setSelId] = useState<string | null>(null);
  const { data: all, isLoading } = useEvents('all');
  const patch = useEventPatch();

  const counts = useMemo(() => {
    const c: Record<string, number> = {};
    (all || []).forEach((e) => { c[e.status] = (c[e.status] || 0) + 1; });
    return c;
  }, [all]);

  const list = useMemo(() => (all || []).filter((e) => e.status === tab), [all, tab]);
  const sel = useMemo(() => (all || []).find((e) => e.id === selId) || null, [all, selId]);

  return (
    <>
      <Topbar title="EVENT 접수" />
      <div className="min-h-screen bg-neutral-50 px-4 md:px-6 py-4 space-y-4 overflow-x-hidden">
        {/* 안내 */}
        <div className="flex items-start gap-2 rounded-xl bg-indigo-50 border border-indigo-100 px-3 py-2.5 text-xs text-indigo-700">
          <Zap size={14} className="shrink-0 mt-0.5" />
          <p>온라인 사전접수 전용입니다. 매장 방문 즉시구매는 <button onClick={() => router.push('/sales/new')} className="underline font-semibold">판매입력</button>에서 EVENT 품목을 선택하세요.</p>
        </div>

        {/* 탭 */}
        <div className="flex gap-1 overflow-x-auto scrollbar-hide border-b border-neutral-200">
          {TABS.map((t) => (
            <button
              key={t.key}
              onClick={() => setTab(t.key as EventStatus)}
              className={`shrink-0 whitespace-nowrap px-3 py-2.5 text-sm font-semibold border-b-2 transition ${
                tab === t.key ? 'border-neutral-900 text-neutral-900' : 'border-transparent text-neutral-400 hover:text-neutral-600'
              }`}
            >
              {t.label}
              {counts[t.key] > 0 && <span className="ml-1.5 text-xs px-1.5 py-0.5 rounded-full bg-neutral-200 text-neutral-600">{counts[t.key]}</span>}
            </button>
          ))}
        </div>

        {/* 리스트 */}
        {isLoading ? (
          <div className="py-16 text-center text-sm text-neutral-400"><Loader2 size={20} className="animate-spin inline" /></div>
        ) : list.length === 0 ? (
          <div className="py-16 text-center text-sm text-neutral-400">{EVENT_STATUS_LABEL[tab]} 건이 없습니다</div>
        ) : (
          <div className="space-y-2">
            {list.map((e) => (
              <button
                key={e.id}
                onClick={() => setSelId(e.id)}
                className="w-full text-left bg-white rounded-xl border border-neutral-200 px-4 py-3 hover:border-neutral-400 transition"
              >
                <div className="flex items-center justify-between gap-2">
                  <div className="min-w-0">
                    <div className="text-[10px] font-mono text-neutral-400">{e.event_number}</div>
                    <div className="text-sm font-bold text-neutral-900 truncate">{e.customer_name}</div>
                  </div>
                  <div className="shrink-0 text-right">
                    <div className="text-sm font-bold text-neutral-900">{won(e.total_amount)}</div>
                    <div className="text-[11px] text-neutral-400 flex items-center gap-0.5 justify-end">
                      {e.receive_method === 'visit' ? <><Store size={11} />매장</> : <><Truck size={11} />택배</>}
                    </div>
                  </div>
                </div>
                <div className="mt-1.5 text-xs text-neutral-500 truncate">
                  {(e.items || []).map((it) => `${it.product_name}${it.slicing ? '(슬라이싱)' : ''}×${it.qty}`).join(', ')}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>

      {/* 상세 패널 */}
      <SlidePanel open={!!sel} onClose={() => setSelId(null)} title="EVENT 접수 상세" className="sm:w-[440px]">
        {sel && <EventDetail ev={sel} patch={patch} onDone={() => setSelId(null)} goSales={() => router.push('/sales')} />}
      </SlidePanel>
    </>
  );
}

function EventDetail({ ev, patch, onDone, goSales }: {
  ev: EventSubmission;
  patch: ReturnType<typeof useEventPatch>;
  onDone: () => void;
  goSales: () => void;
}) {
  const Row = ({ label, value }: { label: string; value: React.ReactNode }) => (
    <div className="flex justify-between gap-2 py-1.5 border-b border-neutral-50 text-sm">
      <span className="text-neutral-400 shrink-0 whitespace-nowrap">{label}</span>
      <span className="text-right text-neutral-800 min-w-0">{value}</span>
    </div>
  );

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
            <span className="text-neutral-700 min-w-0 truncate">{it.product_name}{it.slicing ? ' (슬라이싱+2만)' : ''} ×{it.qty}</span>
            <span className="shrink-0 text-neutral-900">{won(it.unit_price * it.qty)}</span>
          </div>
        ))}
        {ev.slicing_addon > 0 && (
          <div className="flex justify-between text-sm text-neutral-500">
            <span>슬라이싱 가공</span><span>+{won(ev.slicing_addon)}</span>
          </div>
        )}
        <div className="flex justify-between pt-1 mt-1 border-t border-neutral-200 text-sm font-bold">
          <span>합계</span><span>{won(ev.total_amount)}</span>
        </div>
      </div>

      <div>
        <Row label="수령방법" value={ev.receive_method === 'visit' ? '매장 방문' : '택배 발송'} />
        {ev.receive_method !== 'visit' && <Row label="배송지" value={[ev.address1, ev.address2].filter(Boolean).join(' ') || '-'} />}
        {ev.memo && <Row label="메모" value={ev.memo} />}
      </div>

      {/* 액션 */}
      <div className="space-y-2 pt-2">
        {ev.status === 'received' && (
          <button
            disabled={patch.isPending}
            onClick={() => patch.mutate({ id: ev.id, action: 'payment_notice' }, { onSuccess: onDone })}
            className="w-full py-2.5 rounded-lg bg-neutral-900 text-white text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
          >
            {patch.isPending ? <Loader2 size={16} className="animate-spin" /> : <Package size={16} />}
            입금안내 발송 (재고 확인 후)
          </button>
        )}
        {ev.status === 'payment_noticed' && (
          <button
            disabled={patch.isPending}
            onClick={() => {
              if (!window.confirm(`${ev.customer_name}님 입금을 확인하고 판매로 전환합니다. (재고가 차감됩니다)`)) return;
              patch.mutate({ id: ev.id, action: 'confirm_payment' }, { onSuccess: onDone });
            }}
            className="w-full py-2.5 rounded-lg bg-emerald-600 text-white text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
          >
            {patch.isPending ? <Loader2 size={16} className="animate-spin" /> : '입금확인 → 판매 전환'}
          </button>
        )}
        {ev.status === 'converted' && (
          <button onClick={goSales} className="w-full py-2.5 rounded-lg border border-neutral-300 text-sm font-semibold text-neutral-700">
            판매로 전환됨 — 판매관리에서 발송/수령 처리 →
          </button>
        )}
        {ev.status !== 'cancelled' && ev.status !== 'converted' && (
          <button
            disabled={patch.isPending}
            onClick={() => {
              if (!window.confirm('이 접수를 취소합니다.')) return;
              patch.mutate({ id: ev.id, action: 'cancel' }, { onSuccess: onDone });
            }}
            className="w-full py-2 rounded-lg text-xs text-neutral-400 hover:text-red-500"
          >
            접수 취소
          </button>
        )}
      </div>
    </div>
  );
}
