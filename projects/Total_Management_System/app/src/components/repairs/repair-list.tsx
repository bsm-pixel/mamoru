'use client';

import { useState, useMemo, useEffect } from 'react';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from './repair-status-badge';
import { DataGrid, type GridColumn } from '@/components/ui/data-grid';
import { useRepairTabData } from '@/hooks/use-repair-tabs';
import { useUpdateRepairStatus, useUpdateRepairFields, useShipRepair } from '@/hooks/use-repairs';
import { formatKRW, formatPhone, formatDate, formatDateTime } from '@/lib/utils/format';
import {
  Search, Scissors, Package, MapPin, CheckCircle,
  CreditCard, Truck, ClipboardCheck, ClipboardList, Printer,
} from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { RepairPrepSheetModal } from './repair-prep-sheet-modal';
import type { Repair } from '@/lib/supabase/types';
import type { RepairTabKey } from './repair-tab-bar';

interface RepairListProps {
  onSelect?: (id: string) => void;
  selectedId?: string | null;
  initialTab?: RepairTabKey;
  unpaidOnly?: boolean;
  staleOnly?: boolean;
  onClearFilter?: () => void;
  /** PC 그리드(밀집 표) 모드 — 카드 대신 표로 렌더 (PC에서만 전달) */
  gridMode?: boolean;
}

/** 경과일 계산 */
function getDaysElapsed(dateStr: string) {
  return Math.floor((Date.now() - new Date(dateStr).getTime()) / 86400000);
}

/** 날짜 2행 표기 — 위 M/d, 아래 HH:mm (칸 활용). 값 없으면 — */
function TwoLineDate({ iso }: { iso: string | null | undefined }) {
  if (!iso) return <span className="text-neutral-300">—</span>;
  return (
    <div className="leading-tight tabular-nums whitespace-nowrap">
      <div className="text-[13px] text-neutral-700">{formatDate(iso, 'M/d')}</div>
      <div className="text-[10px] text-neutral-400">{formatDate(iso, 'HH:mm')}</div>
    </div>
  );
}

/** 진행방법 배지 */
function ProceedBadge({ type }: { type: string | null }) {
  const t = type || '직접발송';
  const cls = t === '방문수거' ? 'bg-purple-50 text-purple-700'
    : t === '직접방문' ? 'bg-emerald-50 text-emerald-700'
    : 'bg-neutral-100 text-neutral-600'; // 직접발송
  const label = t === '직접방문' ? '직접방문' : t;
  return <span className={`inline-block px-2 py-0.5 rounded-full text-[11px] font-medium whitespace-nowrap ${cls}`}>{label}</span>;
}

/** 날짜(시간없음) + 아래 라벨/시간 — DATE 필드용 (수거요청일/방문예약일) */
function DateLine({ dateStr, sub }: { dateStr: string; sub?: string }) {
  return (
    <div className="leading-tight tabular-nums whitespace-nowrap">
      <div className="text-[13px] text-neutral-700">{formatDate(dateStr, 'M/d')}</div>
      {sub && <div className="text-[10px] text-neutral-400">{sub}</div>}
    </div>
  );
}

/** 수거요청일 칸 — 방문수거=수거요청일 / 직접방문=방문예약(날짜+시간) / 직접발송=라벨 */
function IntakeCell({ repair: r }: { repair: Repair }) {
  if (r.proceed_type === '방문수거') {
    return r.pickup_date ? <DateLine dateStr={r.pickup_date} sub="수거" /> : <span className="text-neutral-300">—</span>;
  }
  if (r.proceed_type === '직접방문') {
    const vd = (r as { visit_date?: string | null; visit_time?: string | null }).visit_date;
    const vt = (r as { visit_time?: string | null }).visit_time;
    return vd ? <DateLine dateStr={vd} sub={vt ? String(vt).slice(0, 5) : '방문'} /> : <span className="text-[11px] text-emerald-600">매장방문</span>;
  }
  return <span className="text-[11px] text-neutral-500">직접발송</span>;
}

/** 검색 필터 (이름/전화/접수번호) */
function matchSearch(r: Repair, q: string): boolean {
  if (!q) return true;
  const lower = q.toLowerCase();
  return (
    (r.name || '').toLowerCase().includes(lower) ||
    (r.phone || '').includes(q) ||
    (r.as_id || '').toLowerCase().includes(lower)
  );
}

export function RepairList({ onSelect, selectedId, initialTab, unpaidOnly, staleOnly, onClearFilter, gridMode }: RepairListProps = {}) {
  const [activeTab, setActiveTab] = useState<RepairTabKey>(initialTab || 'intake');
  const [search, setSearch] = useState('');
  const { tabs, tabData, isLoading } = useRepairTabData();
  // 준비표(트레이형) 다중선택
  const [prepMode, setPrepMode] = useState(false);
  const [checkedIds, setCheckedIds] = useState<Set<string>>(new Set());
  const [showPrep, setShowPrep] = useState(false);
  const toggleCheck = (id: string) => setCheckedIds((prev) => {
    const next = new Set(prev); if (next.has(id)) next.delete(id); else next.add(id); return next;
  });
  const rowClick = (id: string) => { if (prepMode) toggleCheck(id); else onSelect?.(id); };

  // initialTab 변경 시 탭 전환
  useEffect(() => {
    if (initialTab) setActiveTab(initialTab);
  }, [initialTab]);

  // 검색 + 교차 필터 적용
  const filteredRepairs = useMemo(() => {
    let list = tabData[activeTab] || [];
    if (search) list = list.filter((r) => matchSearch(r, search));
    if (unpaidOnly) list = list.filter((r) => !r.paid_at);
    if (staleOnly) {
      const threeDaysAgo = Date.now() - 3 * 86400000;
      list = list.filter((r) => new Date(r.updated_at).getTime() < threeDaysAgo);
    }
    return list;
  }, [tabData, activeTab, search, unpaidOnly, staleOnly]);

  // PC 그리드 컬럼 — 고객(맨앞)·진행방법·접수일·수거요청일·입고일·수량·금액·상태·액션. 날짜는 2행(M/d · HH:mm)
  const gridColumns = useMemo<GridColumn<Repair>[]>(() => [
    { key: 'customer', label: '고객', render: (r) => (
      <div className="min-w-0">
        <div className={`font-bold text-sm truncate ${r.status === 'cancelled' ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>{r.name}</div>
        <div className="text-[11px] text-neutral-400">{formatPhone(r.phone)}</div>
      </div>
    ) },
    { key: 'proceed', label: '진행방법', render: (r) => <ProceedBadge type={r.proceed_type} /> },
    { key: 'received', label: '접수일', render: (r) => <TwoLineDate iso={r.received_at} /> },
    { key: 'pickup', label: '수거요청일', render: (r) => <IntakeCell repair={r} /> },
    { key: 'inbound', label: '입고일', render: (r) => <TwoLineDate iso={(r as { inbound_at?: string | null }).inbound_at ?? null} /> },
    { key: 'qty', label: '수량', render: (r) => (
      <span className="text-xs text-neutral-600 whitespace-nowrap">
        {r.qty_mamoru > 0 && <span>마모루 {r.qty_mamoru}</span>}
        {r.qty_mamoru > 0 && r.qty_other > 0 && <span className="text-neutral-300"> · </span>}
        {r.qty_other > 0 && <span>타사 {r.qty_other}</span>}
        {r.qty_mamoru === 0 && r.qty_other === 0 && <span className="text-neutral-300">—</span>}
      </span>
    ) },
    { key: 'amount', label: '금액', align: 'right', render: (r) => (
      <div className="text-right whitespace-nowrap">
        <span className="font-bold tabular-nums text-indigo-black">{r.total_amount > 0 ? formatKRW(r.total_amount) : '—'}</span>
        {!r.paid_at && r.total_amount > 0 && <div className="text-[10px] text-red-500 font-semibold">미입금</div>}
      </div>
    ) },
    { key: 'status', label: '상태', render: (r) => <RepairStatusBadge status={r.status} proceedType={r.proceed_type} /> },
    { key: 'action', label: '', align: 'right', render: (r) => <InlineAction repair={r} tab={activeTab} /> },
  ], [activeTab]);

  return (
    <div className="space-y-3">
      {/* 검색 */}
      <div className="relative">
        <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
        <input
          type="text"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="이름, 전화번호, 접수번호 검색..."
          className="w-full h-9 pl-9 pr-3 rounded-xl border border-stone-200 bg-white text-sm text-stone-800 placeholder:text-stone-400 focus:outline-none focus:border-stone-400 transition"
        />
      </div>

      {/* 교차 필터 활성 안내 */}
      {(unpaidOnly || staleOnly) && (
        <div className="flex items-center justify-between px-3 py-1.5 rounded-lg bg-red-50 border border-red-200">
          <span className="text-xs font-semibold text-red-700">
            {unpaidOnly ? '미입금 건만 표시 중' : '3일 경과 건만 표시 중'}
          </span>
          {onClearFilter && (
            <button onClick={onClearFilter} className="text-xs text-red-500 hover:text-red-700 font-medium">해제</button>
          )}
        </div>
      )}

      {/* 6탭 파이프라인 + 카운트 뱃지 */}
      <div className="flex gap-1 overflow-x-auto border-b border-stone-200 scrollbar-hide">
        {tabs.map((tab) => {
          const isActive = activeTab === tab.key;
          return (
            <button
              key={tab.key}
              onClick={() => setActiveTab(tab.key)}
              className={`shrink-0 flex items-center gap-1.5 px-3 py-2.5 text-sm font-semibold border-b-2 transition whitespace-nowrap ${
                isActive
                  ? 'border-stone-900 text-stone-900'
                  : 'border-transparent text-stone-500 hover:text-stone-700'
              }`}
            >
              {tab.label}
              {tab.count > 0 && (
                <span className={`inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full text-[10px] font-bold leading-none ${
                  isActive ? 'bg-stone-900 text-white' : 'bg-stone-100 text-stone-600'
                }`}>
                  {tab.count}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* 건수 + 준비표(트레이형) 뽑기 */}
      <div className="flex items-center gap-2 flex-wrap">
        <p className="text-xs text-neutral-500">{filteredRepairs.length}건</p>
        <div className="ml-auto flex items-center gap-2">
          <button
            onClick={() => { setPrepMode(!prepMode); setCheckedIds(new Set()); }}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition ${
              prepMode ? 'bg-neutral-900 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
            }`}
          >
            <ClipboardList size={14} /> {prepMode ? '선택 취소' : '준비표 뽑기'}
          </button>
          {prepMode && checkedIds.size > 0 && (
            <button
              onClick={() => setShowPrep(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-blue-600 text-white text-xs font-medium hover:bg-blue-700 transition"
            >
              <Printer size={14} /> 준비표 인쇄 ({checkedIds.size}건)
            </button>
          )}
        </div>
      </div>
      {prepMode && <p className="text-[11px] text-neutral-400 -mt-1">체크한 건이 A4 1장에 2개씩 (가운데 절취) 나옵니다. 행을 눌러 선택하세요.</p>}

      {/* 목록 */}
      {isLoading ? (
        <div className="space-y-3">
          {Array.from({ length: 5 }).map((_, i) => (
            <Skeleton key={i} className="h-20 w-full rounded-xl" />
          ))}
        </div>
      ) : filteredRepairs.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-16 text-neutral-400">
          <Scissors size={28} className="mb-2 opacity-40" />
          <p className="text-sm">{search ? '검색 결과가 없습니다' : '해당 건이 없습니다'}</p>
        </div>
      ) : gridMode ? (
        <Card padding={false}>
          <DataGrid
            columns={prepMode ? [{
              key: '_chk', label: '', render: (r: Repair) => (
                <input type="checkbox" readOnly checked={checkedIds.has(r.id)} className="w-4 h-4 pointer-events-none" />
              ),
            } as GridColumn<Repair>, ...gridColumns] : gridColumns}
            rows={filteredRepairs}
            getRowKey={(r) => r.id}
            selectedKey={prepMode ? undefined : (selectedId ?? undefined)}
            onSelect={(r) => rowClick(r.id)}
            rowClassName={(r) => `${r.status === 'cancelled' ? 'opacity-60' : ''} ${prepMode && checkedIds.has(r.id) ? 'bg-blue-50' : ''}`}
          />
        </Card>
      ) : (
        <div className="space-y-2">
          {filteredRepairs.map((r) => (
            <RepairCard
              key={r.id}
              repair={r}
              tab={activeTab}
              isSelected={prepMode ? checkedIds.has(r.id) : selectedId === r.id}
              prepMode={prepMode}
              checked={checkedIds.has(r.id)}
              onSelect={rowClick}
            />
          ))}
        </div>
      )}

      {showPrep && <RepairPrepSheetModal repairIds={[...checkedIds]} onClose={() => setShowPrep(false)} />}
    </div>
  );
}

// ── 탭별 특화 카드 ──

interface RepairCardProps {
  repair: Repair;
  tab: RepairTabKey;
  isSelected: boolean;
  onSelect?: (id: string) => void;
  prepMode?: boolean;
  checked?: boolean;
}

function RepairCard({ repair: r, tab, isSelected, onSelect, prepMode = false, checked = false }: RepairCardProps) {
  const days = getDaysElapsed(r.received_at);
  const isCancelled = r.status === 'cancelled';
  const isCompleted = r.status === 'completed';

  // 카드 좌측 border 색상
  const isUnpaid = !r.paid_at && ['cost_notified', 'repairing', 'ready_to_ship', 'shipped'].includes(r.status);
  const borderClass = isCancelled
    ? 'border-l-2 border-l-neutral-300'
    : isCompleted
    ? 'border-l-2 border-l-green-400'
    : (isUnpaid && days >= 3)
    ? 'border-l-2 border-l-red-400'
    : isUnpaid
    ? 'border-l-2 border-l-orange-400'
    : '';

  return (
    <div
      onClick={() => onSelect?.(r.id)}
      className="cursor-pointer"
    >
      <Card className={`hover:bg-neutral-50 transition ${isSelected ? (prepMode ? 'ring-2 ring-blue-500 bg-blue-50' : 'ring-2 ring-terracotta bg-terracotta/5') : ''} ${borderClass} ${isCancelled ? 'opacity-60' : ''}`}>
        <div className="flex items-start justify-between gap-2">
          {prepMode && (
            <input type="checkbox" readOnly checked={checked} className="w-4 h-4 mt-1 shrink-0 pointer-events-none" />
          )}
          <div className="flex-1 min-w-0">
            {/* 상단: 접수일 + 상태 + 진행방식 */}
            <div className="flex items-center gap-2 mb-1 flex-wrap">
              <span className="text-xs text-neutral-400">
                {formatDate(r.received_at, 'M/d HH:mm')}
              </span>
              <RepairStatusBadge status={r.status} proceedType={r.proceed_type} />
              {r.proceed_type && (
                <span className="px-1.5 py-0.5 rounded-full text-[10px] font-medium bg-info-soft text-info">
                  {r.proceed_type}
                </span>
              )}
            </div>

            {/* 고객 정보 */}
            <p className={`text-sm font-semibold truncate ${isCancelled ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>
              {r.name}
            </p>
            <p className="text-xs text-neutral-500">{formatPhone(r.phone)}</p>

            {/* 가위 수량 + 금액 */}
            <div className="flex items-center gap-3 mt-1 text-xs text-neutral-600">
              {r.qty_mamoru > 0 && <span>마모루 {r.qty_mamoru}자루</span>}
              {r.qty_other > 0 && <span>타사 {r.qty_other}자루</span>}
              {r.total_amount > 0 && (
                <span className="font-medium text-terracotta-deep">{formatKRW(r.total_amount)}</span>
              )}
            </div>

            {/* 탭별 특화 정보 */}
            <TabSpecificInfo repair={r} tab={tab} />
          </div>

          {/* 우측: 경과일 + 인라인 액션 */}
          <div className="text-right shrink-0 flex flex-col items-end gap-1.5">
            {days > 0 && !isCompleted && !isCancelled && (
              <p className={`text-[11px] ${days >= 7 ? 'text-error font-medium' : days >= 3 ? 'text-orange-500' : 'text-neutral-400'}`}>
                {days}일 경과
              </p>
            )}
            {!prepMode && <InlineAction repair={r} tab={tab} />}
          </div>
        </div>
      </Card>
    </div>
  );
}

// ── 탭별 특화 정보 ──

function TabSpecificInfo({ repair: r, tab }: { repair: Repair; tab: RepairTabKey }) {
  switch (tab) {
    case 'pickup_needed':
      // 주소 표시
      return r.address ? (
        <div className="flex items-center gap-1 mt-1.5 text-xs text-neutral-500">
          <MapPin size={11} className="shrink-0" />
          <span className="truncate">{r.address}</span>
        </div>
      ) : null;

    case 'in_progress':
      // 입금 상태 + 검수 여부 칩
      return (
        <div className="flex items-center gap-2 mt-1.5">
          {r.paid_at ? (
            <Badge className="bg-green-100 text-green-700 text-[10px]">
              <CheckCircle size={10} className="mr-0.5" />입금완료
            </Badge>
          ) : (
            <Badge className="bg-red-100 text-red-700 text-[10px]">
              <CreditCard size={10} className="mr-0.5" />미입금
            </Badge>
          )}
        </div>
      );

    case 'ready_to_ship':
      // 송장 여부 + 포장 여부
      return (
        <div className="flex items-center gap-2 mt-1.5">
          {r.invoice_number ? (
            <Badge className="bg-green-100 text-green-700 text-[10px]">
              <Package size={10} className="mr-0.5" />{r.invoice_number}
            </Badge>
          ) : (
            <Badge className="bg-neutral-100 text-neutral-500 text-[10px]">송장 미생성</Badge>
          )}
          {r.packed_at && (
            <Badge className="bg-blue-100 text-blue-700 text-[10px]">
              <ClipboardCheck size={10} className="mr-0.5" />포장완료
            </Badge>
          )}
        </div>
      );

    case 'shipped':
      // 송장 + 출고일
      return (
        <div className="flex items-center gap-2 mt-1.5 text-xs text-neutral-500">
          {r.invoice_number && (
            <span className="flex items-center gap-1">
              <Package size={11} /> {r.invoice_number}
            </span>
          )}
          {r.shipped_at && (
            <span>{formatDate(r.shipped_at, 'M/d')} 출고</span>
          )}
          {r.status === 'completed' && (
            <Badge className="bg-green-100 text-green-700 text-[10px]">완료</Badge>
          )}
          {r.status === 'delivered' && (
            <Badge className="bg-blue-100 text-blue-700 text-[10px]">배송완료</Badge>
          )}
          {/* 출고/완료됐는데 미입금 — 사각지대 방지(2026-06-11): 입금확인 누락 표시 */}
          {!r.paid_at && r.total_amount > 0 && (
            <Badge className="bg-red-100 text-red-700 text-[10px]">
              <CreditCard size={10} className="mr-0.5" />미입금
            </Badge>
          )}
        </div>
      );

    default:
      return null;
  }
}

// ── 인라인 퀵 액션 ──

function InlineAction({ repair: r, tab }: { repair: Repair; tab: RepairTabKey }) {
  const updateStatus = useUpdateRepairStatus();
  const updateFields = useUpdateRepairFields();
  const shipRepair = useShipRepair();
  const [showConfirm, setShowConfirm] = useState(false);
  const busy = updateStatus.isPending || updateFields.isPending || shipRepair.isPending;

  const handleClick = (e: React.MouseEvent) => {
    e.stopPropagation();
    setShowConfirm(true);
  };

  // 확인 후 실행할 액션
  const handleConfirm = async () => {
    switch (tab) {
      case 'intake':
        await updateFields.mutateAsync({ id: r.id, confirmed_at: new Date().toISOString() });
        break;
      case 'pickup_needed':
        await updateStatus.mutateAsync({ id: r.id, status: 'pickup_scheduled', note: '수거접수 완료' });
        break;
      case 'in_progress':
        await updateFields.mutateAsync({ id: r.id, paid_at: new Date().toISOString() });
        break;
      case 'ready_to_ship':
        if (!r.invoice_number) {
          await shipRepair.mutateAsync({ id: r.id });
        } else {
          await updateStatus.mutateAsync({ id: r.id, status: 'shipped', shipped_at: new Date().toISOString(), note: '출고완료' });
        }
        break;
    }
  };

  const labels: Record<string, { btn: string; title: string; msg: string }> = {
    intake: { btn: '접수확인', title: '접수 확인', msg: `${r.name}님의 접수를 확인합니다.` },
    pickup_needed: { btn: '수거접수', title: '수거접수 완료', msg: `${r.name}님의 수거접수를 완료합니다.` },
    in_progress: { btn: '입금확인', title: '입금 확인', msg: `${r.name}님의 입금을 확인합니다.` },
    ready_to_ship: !r.invoice_number
      ? { btn: '송장생성', title: '송장 생성', msg: '롯데택배 송장을 생성합니다.' }
      : { btn: '출고완료', title: '출고 완료', msg: `송장 ${r.invoice_number}으로 출고합니다. 알림톡이 발송됩니다.` },
  };

  const label = labels[tab];
  if (!label) return null;
  if (tab === 'in_progress' && r.paid_at) return null;

  const btnColors: Record<string, string> = {
    intake: 'bg-blue-100 text-blue-700 hover:bg-blue-200',
    pickup_needed: 'bg-purple-100 text-purple-700 hover:bg-purple-200',
    in_progress: 'bg-green-100 text-green-700 hover:bg-green-200',
    ready_to_ship: !r.invoice_number ? 'bg-amber-100 text-amber-700 hover:bg-amber-200' : 'bg-green-100 text-green-700 hover:bg-green-200',
  };

  return (
    <>
      <button
        onClick={handleClick}
        disabled={busy}
        className={`px-2 py-1 rounded-md text-[11px] font-semibold transition disabled:opacity-50 ${btnColors[tab]}`}
      >
        {tab === 'ready_to_ship' && !r.invoice_number && <Truck size={11} className="inline mr-0.5" />}
        {label.btn}
      </button>
      {showConfirm && (
        <div onClick={(e) => e.stopPropagation()}>
          <ConfirmModal
            open={showConfirm}
            onClose={() => setShowConfirm(false)}
            onConfirm={handleConfirm}
            title={label.title}
            message={label.msg}
            confirmLabel={label.btn}
          />
        </div>
      )}
    </>
  );
}
