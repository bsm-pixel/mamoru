'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Skeleton } from '@/components/ui/skeleton';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useIsLg } from '@/hooks/use-grid-mode';
import { useReturns, useUpdateReturn } from '@/hooks/use-returns';
import { RETURN_STATUS_LABEL, RETURN_STATUS_COLOR, RETURN_ACTION_LABEL, getAllowedReturnTransitions } from '@/lib/returns/transitions';
import { formatDate, formatPhone } from '@/lib/utils/format';
import { Undo2, Package, Truck } from 'lucide-react';
import type { ReturnRow, ReturnStatus } from '@/lib/supabase/types';

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
  const allowed = getAllowedReturnTransitions(r.status);

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

      {/* 상태 타임라인 */}
      <Card>
        <p className="text-xs font-semibold text-neutral-500 mb-2">진행</p>
        <div className="space-y-1 text-xs text-neutral-500">
          {r.requested_at && <p>수거접수 {formatDate(r.requested_at)}</p>}
          {r.pickup_scheduled_at && <p>수거예약 {formatDate(r.pickup_scheduled_at)}</p>}
          {r.inbound_at && <p className="text-purple-600 font-medium">입고완료 {formatDate(r.inbound_at)} → 반품창고</p>}
          {r.inspected_at && <p>검수완료 {formatDate(r.inspected_at)}</p>}
          {r.completed_at && <p className="text-emerald-600 font-medium">완료 {formatDate(r.completed_at)}</p>}
          {r.cancelled_at && <p className="text-red-500">취소 {formatDate(r.cancelled_at)}</p>}
        </div>
      </Card>

      {/* 상태 전이 액션 */}
      {allowed.length > 0 && (
        <Card>
          <p className="text-xs font-semibold text-neutral-500 mb-2">다음 단계</p>
          <div className="flex flex-wrap gap-2">
            {allowed.map((next) => (
              <button key={next} disabled={update.isPending}
                onClick={() => update.mutate({ id: r.id, status: next })}
                className={`px-3 py-2 rounded-lg text-xs font-semibold transition disabled:opacity-50 ${next === 'cancelled' ? 'border border-red-200 text-red-600 hover:bg-red-50' : 'bg-stone-900 text-white hover:bg-stone-800'}`}>
                {RETURN_ACTION_LABEL[next as ReturnStatus]}
              </button>
            ))}
          </div>
          {allowed.includes('inbound') && <p className="text-[11px] text-neutral-400 mt-2">입고완료 시 구 제품 시리얼이 반품창고로 확정됩니다.</p>}
        </Card>
      )}
    </div>
  );
}
