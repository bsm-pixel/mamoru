'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Skeleton } from '@/components/ui/skeleton';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useIsLg } from '@/hooks/use-grid-mode';
import { useReturns, useUpdateReturn, useShipReturn, useBookReturnPickup } from '@/hooks/use-returns';
import { RETURN_STATUS_LABEL, RETURN_STATUS_COLOR, RETURN_ACTION_LABEL, RETURN_STATUS_HINT, RETURN_PRIMARY_NEXT, getAllowedReturnTransitions } from '@/lib/returns/transitions';
import { StatusStepper } from '@/components/ui/status-stepper';
import { formatDate, formatPhone } from '@/lib/utils/format';
import { Undo2, Package, Truck } from 'lucide-react';
import type { ReturnRow } from '@/lib/supabase/types';

const STATUS_TABS: { value: string; label: string }[] = [
  { value: 'all', label: '전체' },
  { value: 'requested', label: '수거접수' },
  { value: 'pickup_scheduled', label: '수거예약' },
  { value: 'inbound', label: '입고완료' },
  { value: 'inspected', label: '검수완료' },
  { value: 'completed', label: '완료' },
  { value: 'cancelled', label: '취소' },
];

export default function ReturnsPage() {
  const [status, setStatus] = useState('all');
  const [search, setSearch] = useState('');
  const [selectedId, setSelectedId] = useState<string | null>(null);
  // 주문/판매 상세의 [반품·교환에서 관리 →] 링크로 진입 시 검색어 프리필
  useEffect(() => {
    const s = new URLSearchParams(window.location.search).get('search');
    // 마운트 1회 URL 프리필(링크 진입) — 정당한 1회 setState
    // eslint-disable-next-line react-hooks/set-state-in-effect
    if (s) setSearch(s);
  }, []);
  const isLg = useIsLg();
  const { data, isLoading } = useReturns({ status: status === 'all' ? undefined : status, search: search || undefined });
  const returns = data?.returns || [];
  const selected = returns.find((r) => r.id === selectedId) || null;

  const listContent = isLoading ? (
    <div className="p-4 space-y-3">{Array.from({ length: 5 }).map((_, i) => <Skeleton key={i} className="h-16 w-full" />)}</div>
  ) : returns.length === 0 ? (
    <EmptyState icon={Undo2} message="반품·교환수거 건이 없습니다" />
  ) : (
    <div className="divide-y divide-neutral-100">
      {returns.map((r) => (
        <button key={r.id} onClick={() => setSelectedId(r.id)}
          className={`w-full text-left flex items-center gap-3 px-4 py-3 hover:bg-stone-50 transition ${selectedId === r.id ? 'bg-stone-50 border-l-2 border-l-stone-900' : ''}`}>
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2">
              <span className="text-sm font-semibold text-stone-800 truncate">{r.name || '이름없음'}</span>
              <span className={`px-2 py-0.5 rounded text-[10.5px] font-bold ${RETURN_STATUS_COLOR[r.status]}`}>{RETURN_STATUS_LABEL[r.status]}</span>
              <span className="text-[10px] font-semibold text-neutral-400">{r.return_type === 'refund' ? '반품' : '교환'}</span>
            </div>
            <div className="flex items-center gap-2 mt-0.5 text-xs text-neutral-500 min-w-0">
              <span className="font-mono text-[11px] shrink-0">{r.return_number}</span>
              <span className="truncate">{r.product_name}{r.serial_number ? ` · ${r.serial_number}` : ''}</span>
            </div>
          </div>
          {r.pickup_method && <span className="shrink-0 text-[11px] text-neutral-400 flex items-center gap-0.5"><Truck size={11} />{r.pickup_method}</span>}
        </button>
      ))}
    </div>
  );

  return (
    <>
      <Topbar title="반품 · 교환" />
      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 space-y-3">
        <SearchInput value={search} onChange={setSearch} placeholder="반품번호, 이름, 시리얼 검색" />
        <div className="flex gap-1 overflow-x-auto pb-1">
          {STATUS_TABS.map((t) => (
            <button key={t.value} onClick={() => setStatus(t.value)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${status === t.value ? 'bg-stone-900 text-white' : 'bg-stone-100 text-stone-600 hover:bg-stone-200'}`}>
              {t.label}
            </button>
          ))}
        </div>

        {isLg ? (
          <div className="flex gap-4 h-[calc(100vh-220px)]">
            <div className="flex-1 min-w-0 overflow-y-auto"><Card padding={false}>{listContent}</Card></div>
            <div className="w-[400px] shrink-0 overflow-y-auto">
              {selected ? <ReturnDetail r={selected} /> : (
                <div className="flex flex-col items-center justify-center h-60 text-stone-400"><Undo2 size={28} className="mb-2 opacity-40" /><p className="text-xs">목록에서 반품 건을 선택하세요</p></div>
              )}
            </div>
          </div>
        ) : (
          <>
            <Card padding={false}>{listContent}</Card>
            <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="반품 상세">
              {selected && <ReturnDetail r={selected} />}
            </SlidePanel>
          </>
        )}
      </div>
    </>
  );
}

function ReturnDetail({ r }: { r: ReturnRow }) {
  const update = useUpdateReturn();
  const ship = useShipReturn();
  const pickup = useBookReturnPickup();
  const allowed = getAllowedReturnTransitions(r.status);

  // 진행 흐름 — 교환/반품 타입별로 다르게 (교환은 '새 제품 발송'이 핵심 단계)
  const isExchange = r.return_type === 'exchange';
  const shipped = !!r.exchange_out_invoice_number;
  const flowSteps = isExchange
    ? [
        { key: 'requested', label: '접수', at: r.requested_at },
        { key: 'inbound', label: '구제품 회수', at: r.inbound_at },
        { key: 'shipped', label: '새 제품 발송', at: r.exchange_shipped_at },
        { key: 'completed', label: '완료', at: r.completed_at },
      ]
    : [
        { key: 'requested', label: '수거접수', at: r.requested_at },
        { key: 'pickup_scheduled', label: '수거예약', at: r.pickup_scheduled_at },
        { key: 'inbound', label: '입고완료', at: r.inbound_at },
        { key: 'inspected', label: '검수완료', at: r.inspected_at },
        { key: 'completed', label: '완료', at: r.completed_at },
      ];
  const flowCurrent = isExchange
    ? (r.status === 'completed' ? 'completed' : shipped ? 'shipped' : ['inbound', 'inspected'].includes(r.status) ? 'inbound' : 'requested')
    : r.status;

  return (
    <div className="space-y-4">
      <Card>
        <div className="flex items-center justify-between mb-2">
          <h3 className="text-sm font-bold font-mono text-stone-900">{r.return_number}</h3>
          <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${RETURN_STATUS_COLOR[r.status]}`}>{RETURN_STATUS_LABEL[r.status]}</span>
        </div>
        <div className="text-sm space-y-1.5">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-stone-800">{r.name || '이름없음'}</span>
            {r.phone && <a href={`tel:${r.phone}`} className="text-xs text-blue-600">{formatPhone(r.phone)}</a>}
            <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-500">{r.return_type === 'refund' ? '반품환불' : '교환'}</span>
          </div>
          <p className="text-neutral-600 flex items-center gap-1"><Package size={13} className="text-neutral-400" />{r.product_name}{r.serial_number ? ` · ${r.serial_number}` : ''}</p>
          {r.pickup_method && <p className="text-xs text-neutral-500">회수: {r.pickup_method}{r.pickup_date ? ` · 예약 ${formatDate(r.pickup_date)}` : ''}</p>}
          {r.reason && <p className="text-xs text-neutral-400">사유: {r.reason}</p>}
        </div>
      </Card>

      {/* 반품 수거접수 (택배 회수 — 롯데 반품 API ustRtgSctCd=02). 완료·취소 건은 숨김 */}
      {r.pickup_method === '택배수거' && r.status !== 'completed' && r.status !== 'cancelled' && (
        <Card>
          <p className="text-xs font-semibold text-neutral-500 mb-2">반품 수거접수 (고객집 → 매장 회수)</p>
          {r.pickup_invoice_number ? (
            <div className="text-xs text-blue-700 bg-blue-50 rounded-lg px-3 py-2">
              ✓ 수거 송장 <b className="font-mono">{r.pickup_invoice_number}</b>
              {r.pickup_booked_at && <span className="text-neutral-400 ml-1">({formatDate(r.pickup_booked_at)})</span>}
              <p className="text-[11px] text-neutral-400 mt-1">접수 후 취소는 ALPS 화면에서 수동으로 해주세요(롯데 취소 API 미지원).</p>
            </div>
          ) : (
            <>
              <button disabled={pickup.isPending} onClick={() => pickup.mutate(r.id)}
                className="w-full flex items-center justify-center gap-1.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 transition disabled:opacity-50">
                <Truck size={13} /> {pickup.isPending ? '접수 중…' : '롯데 반품 수거접수'}
              </button>
              <p className="text-[11px] text-neutral-400 mt-1.5">고객집으로 롯데 기사가 방문 수거합니다. 접수 후 취소는 ALPS에서 수동.</p>
            </>
          )}
        </Card>
      )}

      {/* 교환 출고 송장 (배송 교환 — 새 제품 발송) */}
      {r.return_type === 'exchange' && r.new_product_name && (
        <Card>
          <p className="text-xs font-semibold text-neutral-500 mb-2">교환 출고 (새 제품 발송)</p>
          <p className="text-sm text-neutral-700 mb-2">
            {r.new_product_name}{r.new_serial_number ? ` · ${r.new_serial_number}` : ''}
          </p>
          {r.exchange_out_invoice_number ? (
            <div className="text-xs text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2">
              ✓ 출고 송장 <b className="font-mono">{r.exchange_out_invoice_number}</b>
              {r.exchange_shipped_at && <span className="text-neutral-400 ml-1">({formatDate(r.exchange_shipped_at)})</span>}
            </div>
          ) : (
            <button disabled={ship.isPending} onClick={() => ship.mutate(r.id)}
              className="w-full flex items-center justify-center gap-1.5 py-2 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition disabled:opacity-50">
              <Truck size={13} /> {ship.isPending ? '발행 중…' : '교환 출고 송장 발행'}
            </button>
          )}
          <p className="text-[11px] text-neutral-400 mt-1.5">새 제품 1개만 담긴 롯데 송장을 발행합니다(원 주문 다품목이어도 교환품만).</p>
        </Card>
      )}

      {/* 상태 타임라인 */}
      {/* 진행 흐름 — 항상 표시(완료·취소 포함) */}
      <Card>
        <p className="text-xs font-semibold text-neutral-500 mb-3">진행 흐름</p>
        <StatusStepper
          steps={flowSteps}
          currentKey={flowCurrent}
          cancelled={r.status === 'cancelled'}
          cancelledAt={r.cancelled_at}
        />
        {/* 지금 할 일 + 액션 (진행 가능한 상태만) */}
        {allowed.length > 0 && (() => {
          const primary = RETURN_PRIMARY_NEXT[r.status];
          const secondary = allowed.filter((s) => s !== primary && s !== 'cancelled');
          return (
            <div className="mt-3 pt-3 border-t border-neutral-100">
              <p className="text-[11px] text-neutral-500 bg-neutral-50 rounded-lg px-3 py-2 mb-2.5 leading-relaxed">
                {RETURN_STATUS_HINT[r.status]}
              </p>
              {primary && allowed.includes(primary) && (
                <button disabled={update.isPending} onClick={() => update.mutate({ id: r.id, status: primary })}
                  className="w-full py-2.5 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 transition disabled:opacity-50 mb-1.5">
                  {RETURN_ACTION_LABEL[primary]} →
                </button>
              )}
              <div className="flex items-center gap-2">
                {secondary.map((next) => (
                  <button key={next} disabled={update.isPending} onClick={() => update.mutate({ id: r.id, status: next })}
                    className="flex-1 py-2 rounded-lg border border-neutral-200 text-xs font-medium text-neutral-600 hover:bg-neutral-50 transition disabled:opacity-50">
                    {RETURN_ACTION_LABEL[next]}
                  </button>
                ))}
                {allowed.includes('cancelled') && (
                  <button disabled={update.isPending} onClick={() => update.mutate({ id: r.id, status: 'cancelled' })}
                    className="px-3 py-2 rounded-lg text-xs font-medium text-red-500 hover:bg-red-50 transition disabled:opacity-50">
                    취소
                  </button>
                )}
              </div>
            </div>
          );
        })()}
      </Card>
    </div>
  );
}
