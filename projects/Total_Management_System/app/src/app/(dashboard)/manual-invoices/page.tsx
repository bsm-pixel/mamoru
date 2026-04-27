'use client';

import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { QuickInvoiceForm } from '@/components/manual-invoices/quick-invoice-form';
import { useTodayManualInvoices, useCancelManualInvoice } from '@/hooks/use-manual-invoices';
import type { ManualInvoice } from '@/lib/supabase/types';
import { Truck, X, Copy } from 'lucide-react';
import toast from 'react-hot-toast';

function formatTime(iso: string): string {
  try {
    return new Date(iso).toLocaleTimeString('ko-KR', {
      timeZone: 'Asia/Seoul',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false,
    });
  } catch {
    return '';
  }
}

async function copyToClipboard(text: string) {
  try {
    await navigator.clipboard.writeText(text);
    toast.success('송장번호 복사 완료');
  } catch {
    toast.error('복사 실패');
  }
}

function TodayInvoiceRow({ invoice, onCancel, isCancelling }: {
  invoice: ManualInvoice;
  onCancel: (id: string) => void;
  isCancelling: boolean;
}) {
  return (
    <div className="flex items-center gap-3 py-2.5 border-b border-neutral-100 last:border-b-0">
      <button
        type="button"
        onClick={() => copyToClipboard(invoice.invoice_number)}
        className="font-mono text-sm font-semibold text-indigo-black hover:text-terracotta transition flex items-center gap-1.5 shrink-0"
        title="송장번호 복사"
      >
        {invoice.invoice_number}
        <Copy size={11} className="opacity-50" />
      </button>
      <div className="flex-1 min-w-0">
        <p className="text-xs text-neutral-700 truncate">
          <span className="font-semibold">{invoice.customer_name}</span>
          <span className="text-neutral-400 mx-1.5">·</span>
          {invoice.goods_name}
        </p>
      </div>
      <span className="text-[11px] text-neutral-400 shrink-0">{formatTime(invoice.created_at)}</span>
      <button
        type="button"
        onClick={() => {
          if (window.confirm(`송장번호 ${invoice.invoice_number}을(를) 정말 취소하시겠습니까?\n\nALPS에 취소 요청이 전송됩니다.`)) {
            onCancel(invoice.id);
          }
        }}
        disabled={isCancelling}
        className="w-7 h-7 rounded-md text-neutral-400 hover:text-red-500 hover:bg-red-50 transition flex items-center justify-center disabled:opacity-30 shrink-0"
        title="송장 취소"
      >
        <X size={14} />
      </button>
    </div>
  );
}

export default function ManualInvoicesPage() {
  const { data, isLoading } = useTodayManualInvoices();
  const cancelInvoice = useCancelManualInvoice();
  const invoices = data?.invoices ?? [];

  return (
    <>
      <Topbar title="빠른 송장" />
      <div className="px-4 md:px-6 py-5 max-w-[800px] mx-auto space-y-5">
        {/* 발급 폼 */}
        <Card>
          <div className="flex items-center gap-2 mb-4 pb-3 border-b border-neutral-100">
            <Truck size={18} className="text-terracotta" />
            <h2 className="text-sm font-bold text-indigo-black">새 송장 발급</h2>
          </div>
          <QuickInvoiceForm />
        </Card>

        {/* 오늘 발급 미니 리스트 */}
        <Card>
          <div className="flex items-center justify-between mb-3 pb-3 border-b border-neutral-100">
            <div className="flex items-center gap-2">
              <h2 className="text-sm font-bold text-indigo-black">오늘 발급한 별도 송장</h2>
              {invoices.length > 0 && (
                <span className="text-xs px-2 py-0.5 rounded-full bg-neutral-100 text-neutral-600 font-semibold">
                  {invoices.length}건
                </span>
              )}
            </div>
            <span className="text-[11px] text-neutral-400">송장번호 클릭 → 복사</span>
          </div>

          {isLoading ? (
            <div className="py-8 text-center text-xs text-neutral-400">불러오는 중...</div>
          ) : invoices.length === 0 ? (
            <div className="py-8 text-center text-xs text-neutral-400">
              오늘 발급한 별도 송장이 없습니다.
            </div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {invoices.map((inv) => (
                <TodayInvoiceRow
                  key={inv.id}
                  invoice={inv}
                  onCancel={(id) => cancelInvoice.mutate({ id })}
                  isCancelling={cancelInvoice.isPending}
                />
              ))}
            </div>
          )}
        </Card>
      </div>
    </>
  );
}
