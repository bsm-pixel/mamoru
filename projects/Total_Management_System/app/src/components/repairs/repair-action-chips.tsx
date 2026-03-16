'use client';

import { useState } from 'react';
import { Button } from '@/components/ui/button';
import {
  useUpdateRepairStatus,
  useUpdateRepairFields,
  useShipRepair,
  useSendRepairNotification,
} from '@/hooks/use-repairs';
import { formatKRW } from '@/lib/utils/format';
import type { Repair } from '@/lib/supabase/types';
import { Check, FileText, CreditCard, Truck, Package } from 'lucide-react';

interface RepairActionChipsProps {
  repair: Repair;
}

/** 진행중 탭 전용: 수리내역서 / 입금확인 / 송장생성 / 포장완료 칩 바 */
export function RepairActionChips({ repair: r }: RepairActionChipsProps) {
  const updateStatus = useUpdateRepairStatus();
  const updateFields = useUpdateRepairFields();
  const shipRepair = useShipRepair();
  const sendNotify = useSendRepairNotification();
  const [showReportConfirm, setShowReportConfirm] = useState(false);

  const isPaid = !!r.paid_at;
  const hasInvoice = !!r.invoice_number;
  const isPacked = !!r.packed_at;

  // 수리내역서 열기 (외부 페이지)
  const handleOpenReport = () => {
    window.open(
      `${window.location.origin}/api/repair/report?id=${r.id}`,
      '_blank'
    );
  };

  // 입금확인
  const handleMarkPaid = () => {
    updateFields.mutate({ id: r.id, paid_at: new Date().toISOString() });
  };

  // 송장생성
  const handleCreateInvoice = () => {
    shipRepair.mutate({ id: r.id });
  };

  // 포장완료
  const handleMarkPacked = () => {
    updateFields.mutate({ id: r.id, packed_at: new Date().toISOString() });
  };

  return (
    <div className="flex flex-wrap gap-1 mt-2">
      {/* 수리내역서 */}
      <button
        onClick={handleOpenReport}
        className="inline-flex items-center gap-0.5 px-2 py-1 rounded-md text-[11px] font-medium bg-neutral-100 text-neutral-600 hover:bg-neutral-200 transition"
      >
        <FileText size={11} />
        내역서
      </button>

      {/* 입금확인 */}
      <button
        onClick={isPaid ? undefined : handleMarkPaid}
        disabled={isPaid || updateFields.isPending}
        className={`inline-flex items-center gap-0.5 px-2 py-1 rounded-md text-[11px] font-medium transition ${
          isPaid
            ? 'bg-green-100 text-green-700'
            : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
        }`}
      >
        {isPaid ? <Check size={11} /> : <CreditCard size={11} />}
        {isPaid ? `입금 ✓` : '입금확인'}
      </button>

      {/* 송장생성 */}
      <button
        onClick={hasInvoice ? undefined : handleCreateInvoice}
        disabled={hasInvoice || shipRepair.isPending}
        className={`inline-flex items-center gap-0.5 px-2 py-1 rounded-md text-[11px] font-medium transition ${
          hasInvoice
            ? 'bg-green-100 text-green-700'
            : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
        }`}
      >
        {hasInvoice ? <Check size={11} /> : <Truck size={11} />}
        {hasInvoice ? `송장 ✓` : '송장생성'}
      </button>

      {/* 포장완료 */}
      <button
        onClick={isPacked ? undefined : handleMarkPacked}
        disabled={isPacked || updateFields.isPending}
        className={`inline-flex items-center gap-0.5 px-2 py-1 rounded-md text-[11px] font-medium transition ${
          isPacked
            ? 'bg-green-100 text-green-700'
            : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
        }`}
      >
        {isPacked ? <Check size={11} /> : <Package size={11} />}
        {isPacked ? '포장 ✓' : '포장완료'}
      </button>
    </div>
  );
}
