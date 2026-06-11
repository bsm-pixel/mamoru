'use client';

import { useState, useMemo, useEffect } from 'react';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from './repair-status-badge';
import { useRepairTabData } from '@/hooks/use-repair-tabs';
import { useUpdateRepairStatus, useUpdateRepairFields, useShipRepair } from '@/hooks/use-repairs';
import { formatKRW, formatPhone, formatDate, formatDateTime } from '@/lib/utils/format';
import {
  Search, Scissors, Package, MapPin, CheckCircle,
  CreditCard, Truck, ClipboardCheck,
} from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import type { Repair } from '@/lib/supabase/types';
import type { RepairTabKey } from './repair-tab-bar';

interface RepairListProps {
  onSelect?: (id: string) => void;
  selectedId?: string | null;
  initialTab?: RepairTabKey;
  unpaidOnly?: boolean;
  staleOnly?: boolean;
  onClearFilter?: () => void;
}

/** 경과일 계산 */
function getDaysElapsed(dateStr: string) {
  return Math.floor((Date.now() - new Date(dateStr).getTime()) / 86400000);
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

export function RepairList({ onSelect, selectedId, initialTab, unpaidOnly, staleOnly, onClearFilter }: RepairListProps = {}) {
  const [activeTab, setActiveTab] = useState<RepairTabKey>(initialTab || 'intake');
  const [search, setSearch] = useState('');
  const { tabs, tabData, isLoading } = useRepairTabData();

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

      {/* 건수 */}
      <p className="text-xs text-neutral-500">{filteredRepairs.length}건</p>

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
      ) : (
        <div className="space-y-2">
          {filteredRepairs.map((r) => (
            <RepairCard
              key={r.id}
              repair={r}
              tab={activeTab}
              isSelected={selectedId === r.id}
              onSelect={onSelect}
            />
          ))}
        </div>
      )}
    </div>
  );
}

// ── 탭별 특화 카드 ──

interface RepairCardProps {
  repair: Repair;
  tab: RepairTabKey;
  isSelected: boolean;
  onSelect?: (id: string) => void;
}

function RepairCard({ repair: r, tab, isSelected, onSelect }: RepairCardProps) {
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
      <Card className={`hover:bg-neutral-50 transition ${isSelected ? 'ring-2 ring-terracotta bg-terracotta/5' : ''} ${borderClass} ${isCancelled ? 'opacity-60' : ''}`}>
        <div className="flex items-start justify-between gap-2">
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
            <InlineAction repair={r} tab={tab} />
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
