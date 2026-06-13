'use client';

import { useCustomer } from '@/hooks/use-customers';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { activityDisplay } from '@/lib/customer/display';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { X, ShoppingBag, FileSignature, MessageSquare, Wrench, Copy, ExternalLink } from 'lucide-react';
import Link from 'next/link';
import toast from 'react-hot-toast';

const CONSULTATION_TYPE_LABEL: Record<string, string> = { store_visit: '매장방문', field_request: '출장', talk_consult: '온라인상담' };

interface Props {
  customerId: string;
  open: boolean;
  onClose: () => void;
}

export function CustomerQuickModal({ customerId, open, onClose }: Props) {
  const { data, isLoading } = useCustomer(customerId);

  if (!open) return null;

  const TIMELINE_ICON = { sale: ShoppingBag, contract: FileSignature, consultation: MessageSquare, repair: Wrench };
  const TIMELINE_COLOR = { sale: 'text-green-600 bg-green-50', contract: 'text-purple-600 bg-purple-50', consultation: 'text-blue-600 bg-blue-50', repair: 'text-orange-600 bg-orange-50' };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-[90vw] max-w-[420px] max-h-[80vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        {/* 헤더 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">고객 정보</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {isLoading ? (
            <div className="space-y-3">
              <Skeleton className="h-16 w-full" />
              <Skeleton className="h-24 w-full" />
            </div>
          ) : !data ? (
            <p className="text-sm text-neutral-400 text-center py-4">고객 정보를 찾을 수 없습니다</p>
          ) : (() => {
            const c = data.customer;
            const { sales, contracts, consultations, repairs, summary } = data as Record<string, unknown> & typeof data;

            // RFM
            const lastSaleDate = (summary as Record<string, unknown>).lastSaleDate as string | null;
            const daysSince = lastSaleDate ? Math.floor((Date.now() - new Date(lastSaleDate).getTime()) / (1000 * 60 * 60 * 24)) : 999;
            const rfm = summary.totalSales >= 3 || summary.totalSalesAmount >= 500000 ? 'VIP' : daysSince > 180 ? '휴면' : '일반';
            const rfmColor = rfm === 'VIP' ? 'bg-amber-100 text-amber-700' : rfm === '휴면' ? 'bg-neutral-200 text-neutral-500' : 'bg-blue-100 text-blue-700';

            // 타임라인 (최근 5건)
            type TItem = { type: string; date: string; title: string; amount?: number; href: string };
            const timeline: TItem[] = [
              ...sales.map((s) => ({ type: 'sale', date: s.sale_date, title: s.sale_number, amount: s.paid_amount, href: `/sales/${s.id}` })),
              ...contracts.map((ct) => ({ type: 'contract', date: ct.created_at, title: ct.contract_number, amount: ct.final_amount, href: `/contracts/${ct.id}` })),
              ...consultations.map((con) => ({ type: 'consultation', date: con.visit_date || con.created_at, title: CONSULTATION_TYPE_LABEL[con.consultation_type] || con.consultation_type, href: `/consultations/${con.id}` })),
              ...((repairs as unknown as Array<{ id: string; repair_number: string; total_cost: number | null; created_at: string }>) || []).map((r) => ({ type: 'repair', date: r.created_at, title: r.repair_number || '복원수리', amount: r.total_cost || 0, href: `/repairs/${r.id}` })),
            ].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 5);

            return (
              <>
                {/* 고객 기본 정보 */}
                <div className="flex items-center gap-3">
                  <div className="flex-1">
                    <div className="flex items-center gap-2">
                      <h4 className="text-base font-bold text-neutral-800">{activityDisplay(c.activity_name, c.name)}</h4>
                      <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${rfmColor}`}>{rfm}</span>
                      {c.customer_type && c.customer_type !== 'retail' && (
                        <Badge className="bg-neutral-100 text-neutral-600">{c.customer_type}</Badge>
                      )}
                    </div>
                    {c.phone && (
                      <button
                        onClick={() => { navigator.clipboard.writeText(c.phone || ''); toast.success('전화번호 복사됨'); }}
                        className="flex items-center gap-1 text-sm text-neutral-500 hover:text-neutral-700 mt-0.5"
                      >
                        {formatPhone(c.phone)}
                        <Copy size={10} />
                      </button>
                    )}
                  </div>
                </div>

                {/* 요약 */}
                <div className="grid grid-cols-3 gap-2">
                  <div className="bg-neutral-50 rounded-lg p-2 text-center">
                    <p className="text-[10px] text-neutral-500">판매</p>
                    <p className="text-sm font-bold">{summary.totalSales}건</p>
                  </div>
                  <div className="bg-neutral-50 rounded-lg p-2 text-center">
                    <p className="text-[10px] text-neutral-500">총액</p>
                    <p className="text-sm font-bold">{formatKRW(summary.totalSalesAmount)}</p>
                  </div>
                  <div className="bg-neutral-50 rounded-lg p-2 text-center">
                    <p className="text-[10px] text-neutral-500">미수금</p>
                    <p className={`text-sm font-bold ${c.outstanding_balance > 0 ? 'text-red-500' : 'text-neutral-400'}`}>
                      {formatKRW(c.outstanding_balance)}
                    </p>
                  </div>
                </div>

                {/* 최근 거래 */}
                {timeline.length > 0 && (
                  <div>
                    <p className="text-xs font-semibold text-neutral-500 mb-2">최근 거래</p>
                    <div className="space-y-1">
                      {timeline.map((item, i) => {
                        const Icon = TIMELINE_ICON[item.type as keyof typeof TIMELINE_ICON];
                        const color = TIMELINE_COLOR[item.type as keyof typeof TIMELINE_COLOR];
                        return (
                          <Link key={i} href={item.href}
                            className="flex items-center gap-2 py-1.5 hover:bg-neutral-50 rounded px-1 transition"
                            onClick={onClose}
                          >
                            <div className={`w-5 h-5 rounded flex items-center justify-center shrink-0 ${color}`}>
                              <Icon size={10} />
                            </div>
                            <span className="text-xs text-neutral-700 flex-1 truncate">{item.title}</span>
                            {item.amount !== undefined && item.amount > 0 && (
                              <span className="text-xs font-medium text-neutral-600">{formatKRW(item.amount)}</span>
                            )}
                            <span className="text-[10px] text-neutral-400">{formatDate(item.date)}</span>
                          </Link>
                        );
                      })}
                    </div>
                  </div>
                )}
              </>
            );
          })()}
        </div>

        {/* 하단 */}
        {data && (
          <div className="px-4 py-3 border-t border-neutral-200">
            <Link
              href={`/customers/${customerId}`}
              onClick={onClose}
              className="flex items-center justify-center gap-2 w-full py-2 rounded-lg bg-neutral-100 text-sm text-neutral-600 hover:bg-neutral-200 transition"
            >
              <ExternalLink size={14} />
              고객 상세 페이지
            </Link>
          </div>
        )}
      </div>
    </div>
  );
}
