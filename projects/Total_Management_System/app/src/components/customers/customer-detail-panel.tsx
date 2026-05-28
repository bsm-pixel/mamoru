'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomer, useUpdateCustomer } from '@/hooks/use-customers';
import { useCustomerManualInvoices } from '@/hooks/use-manual-invoices';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import {
  Save, ShoppingBag, FileSignature, MessageSquare, Wrench,
  Clock, Pencil, X, Truck,
} from 'lucide-react';
import toast from 'react-hot-toast';
import { TagBadges } from '@/components/shared/tag-selector';

const TYPE_OPTIONS = [
  { value: 'retail', label: '일반' },
  { value: 'online', label: '온라인' },
  { value: 'dealer', label: '딜러' },
  { value: 'academy', label: '아카데미' },
];

const TYPE_COLOR: Record<string, string> = {
  retail: 'bg-neutral-100 text-neutral-600',
  online: 'bg-blue-100 text-blue-700',
  dealer: 'bg-purple-100 text-purple-700',
  academy: 'bg-emerald-100 text-emerald-700',
};

const PAYMENT_STATUS_LABEL: Record<string, string> = {
  paid: '결제완료', unpaid: '미결제', partial: '부분결제',
};

const CONSULTATION_TYPE_LABEL: Record<string, string> = {
  store_visit: '매장방문', field_request: '출장', talk_consult: '온라인상담',
};

const REPAIR_STATUS_LABEL: Record<string, string> = {
  intake: '접수', pickup_scheduled: '수거접수', cost_notified: '비용안내',
  repairing: '작업중', shipped: '출고', delivered: '배송완료',
  completed: '완료', cancelled: '취소',
};

interface Props {
  customerId: string;
}

export function CustomerDetailPanel({ customerId }: Props) {
  const router = useRouter();
  const { data, isLoading } = useCustomer(customerId);
  const { data: manualInvoicesData } = useCustomerManualInvoices(customerId);
  const updateCustomer = useUpdateCustomer();
  const [tab, setTab] = useState<'profile' | 'timeline'>('profile');
  const [editing, setEditing] = useState(false);
  const [editMemo, setEditMemo] = useState('');

  if (isLoading) {
    return <div className="space-y-4"><Skeleton className="h-32" /><Skeleton className="h-24" /><Skeleton className="h-48" /></div>;
  }
  if (!data?.customer) {
    return <p className="text-sm text-neutral-400 text-center py-8">고객 정보를 찾을 수 없습니다</p>;
  }

  const { customer: c, sales, contracts, consultations, repairs, summary } = data;
  const manualInvoices = manualInvoicesData?.invoices ?? [];

  // 통합 타임라인 생성 — 시간순 역순
  const timeline = [
    ...sales.map((s) => ({
      type: 'sale' as const,
      id: s.id,
      date: s.sale_date,
      title: s.sale_number,
      sub: PAYMENT_STATUS_LABEL[s.payment_status] || s.payment_status,
      amount: s.paid_amount,
      icon: ShoppingBag,
      color: 'text-blue-600 bg-blue-50',
      cancelled: false,
    })),
    ...contracts.map((ct) => ({
      type: 'contract' as const,
      id: ct.id,
      date: ct.created_at,
      title: ct.contract_number,
      sub: ct.status,
      amount: ct.final_amount,
      icon: FileSignature,
      color: 'text-purple-600 bg-purple-50',
      cancelled: false,
    })),
    ...consultations.map((con) => ({
      type: 'consultation' as const,
      id: con.id,
      date: con.visit_date || con.created_at,
      title: CONSULTATION_TYPE_LABEL[con.consultation_type] || con.consultation_type,
      sub: con.status,
      amount: null as number | null,
      icon: MessageSquare,
      color: 'text-amber-600 bg-amber-50',
      cancelled: false,
    })),
    ...repairs.map((r) => ({
      type: 'repair' as const,
      id: r.id,
      date: r.created_at,
      title: r.repair_number,
      sub: REPAIR_STATUS_LABEL[r.status] || r.status,
      amount: r.total_cost,
      icon: Wrench,
      color: 'text-rose-600 bg-rose-50',
      cancelled: false,
    })),
    ...manualInvoices.map((mi) => ({
      type: 'manual_invoice' as const,
      id: mi.id,
      date: mi.created_at,
      title: mi.goods_name,
      sub: `송장 ${mi.invoice_number}`,
      amount: null as number | null,
      icon: Truck,
      color: 'text-stone-900 bg-stone-100',
      cancelled: !!mi.cancelled_at,
    })),
  ].sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());

  function startMemoEdit() {
    setEditMemo(c.memo || '');
    setEditing(true);
  }

  async function saveMemo() {
    await updateCustomer.mutateAsync({ id: customerId, memo: editMemo });
    setEditing(false);
  }

  const tabs = [
    { key: 'profile' as const, label: '프로필' },
    { key: 'timeline' as const, label: `이력 (${timeline.length})` },
  ];

  return (
    <div className="space-y-4">
      {/* 헤더 */}
      <div className="flex items-center justify-between">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-base font-bold text-stone-900">{c.name}</h2>
            <Badge className={TYPE_COLOR[c.customer_type] || TYPE_COLOR.retail}>
              {TYPE_OPTIONS.find(t => t.value === c.customer_type)?.label || '일반'}
            </Badge>
          </div>
          <p className="text-xs text-neutral-500 mt-0.5">
            {formatPhone(c.phone) || '연락처 없음'}
            {c.company_name && ` · ${c.company_name}`}
          </p>
        </div>
        <Button variant="ghost" size="sm" onClick={() => router.push(`/customers/${customerId}`)}>
          상세 페이지
        </Button>
      </div>

      {/* 요약 카드 */}
      <div className="grid grid-cols-3 gap-2">
        <div className="rounded-lg bg-neutral-50 px-3 py-2">
          <p className="text-[11px] text-neutral-500">판매</p>
          <p className="text-sm font-bold">{summary.totalSales}건</p>
        </div>
        <div className="rounded-lg bg-neutral-50 px-3 py-2">
          <p className="text-[11px] text-neutral-500">매출</p>
          <p className="text-sm font-bold">{formatKRW(summary.totalSalesAmount)}</p>
        </div>
        <div className="rounded-lg bg-neutral-50 px-3 py-2">
          <p className="text-[11px] text-neutral-500">미수금</p>
          <p className={`text-sm font-bold ${c.outstanding_balance > 0 ? 'text-red-500' : 'text-neutral-400'}`}>
            {formatKRW(c.outstanding_balance)}
          </p>
        </div>
      </div>

      {/* 탭 */}
      <div className="flex border-b border-neutral-200">
        {tabs.map(t => (
          <button
            key={t.key}
            onClick={() => setTab(t.key)}
            className={`px-4 py-2 text-xs font-semibold border-b-2 transition ${
              tab === t.key
                ? 'border-indigo-black text-stone-900'
                : 'border-transparent text-neutral-400 hover:text-neutral-600'
            }`}
          >
            {t.label}
          </button>
        ))}
      </div>

      {/* 탭 내용 */}
      {tab === 'profile' ? (
        <div className="space-y-4">
          {/* 기본 정보 */}
          <Card>
            <div className="grid grid-cols-2 gap-3 text-sm">
              <div>
                <span className="text-xs text-neutral-500">연락처</span>
                <p>{formatPhone(c.phone) || '-'}</p>
              </div>
              <div>
                <span className="text-xs text-neutral-500">유형</span>
                <p>{TYPE_OPTIONS.find(t => t.value === c.customer_type)?.label || '일반'}</p>
              </div>
              <div>
                <span className="text-xs text-neutral-500">등록일</span>
                <p>{formatDate(c.created_at)}</p>
              </div>
              <div>
                <span className="text-xs text-neutral-500">매장명 (근무지)</span>
                <p>{c.company_name || '-'}</p>
              </div>
              <div className="col-span-2">
                <span className="text-xs text-neutral-500">주소</span>
                <p>{[c.postcode, c.address_road, c.address_detail].filter(Boolean).join(' ') || '-'}</p>
              </div>
              {(c as unknown as { tags?: string[] }).tags && (c as unknown as { tags?: string[] }).tags!.length > 0 && (
                <div className="col-span-2">
                  <span className="text-xs text-neutral-500 mb-1 block">태그</span>
                  <TagBadges tags={(c as unknown as { tags?: string[] }).tags} />
                </div>
              )}
              {c.email && (
                <div className="col-span-2">
                  <span className="text-xs text-neutral-400">이메일</span>
                  <p className="text-neutral-500 text-xs">{c.email}</p>
                </div>
              )}
            </div>
          </Card>

          {/* 메모 */}
          <Card>
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-xs font-bold text-stone-900">메모</h3>
              {!editing ? (
                <button onClick={startMemoEdit} className="text-neutral-400 hover:text-neutral-600">
                  <Pencil size={14} />
                </button>
              ) : (
                <div className="flex gap-1">
                  <button onClick={() => setEditing(false)} className="text-neutral-400 hover:text-neutral-600">
                    <X size={14} />
                  </button>
                  <button onClick={saveMemo} disabled={updateCustomer.isPending}
                    className="text-blue-600 hover:text-blue-700">
                    <Save size={14} />
                  </button>
                </div>
              )}
            </div>
            {editing ? (
              <textarea
                value={editMemo}
                onChange={(e) => setEditMemo(e.target.value)}
                rows={4}
                placeholder="고객 특이사항, 참고사항 등"
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400 resize-none"
                autoFocus
              />
            ) : (
              <p className="text-sm text-neutral-600 whitespace-pre-wrap">
                {c.memo || '메모 없음'}
              </p>
            )}
          </Card>
        </div>
      ) : (
        /* 통합 타임라인 */
        <div className="space-y-0">
          {timeline.length === 0 ? (
            <p className="text-sm text-neutral-400 text-center py-8">이력이 없습니다</p>
          ) : (
            timeline.map((item, i) => {
              const Icon = item.icon;
              return (
                <div
                  key={`${item.type}-${item.id}`}
                  onClick={() => {
                    if (item.type === 'sale') router.push(`/sales/${item.id}`);
                    else if (item.type === 'contract') router.push(`/contracts/${item.id}`);
                    else if (item.type === 'manual_invoice') router.push('/manual-invoices');
                  }}
                  className={`flex items-start gap-3 py-3 ${
                    i < timeline.length - 1 ? 'border-b border-neutral-100' : ''
                  } ${item.type === 'sale' || item.type === 'contract' || item.type === 'manual_invoice' ? 'cursor-pointer hover:bg-stone-50/40 -mx-1 px-1 rounded' : ''}`}
                >
                  <div className={`w-7 h-7 rounded-full flex items-center justify-center shrink-0 ${item.color}`}>
                    <Icon size={14} />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between">
                      <p className={`text-sm font-medium truncate ${item.cancelled ? 'text-neutral-400 line-through' : 'text-stone-900'}`}>
                        {item.type === 'manual_invoice' && (
                          <Badge className="bg-stone-100 text-stone-900 text-[10px] mr-1.5">빠른송장</Badge>
                        )}
                        {item.title}
                      </p>
                      {item.amount != null && item.amount > 0 && (
                        <span className="text-sm font-bold shrink-0 ml-2">{formatKRW(item.amount)}</span>
                      )}
                    </div>
                    <div className="flex items-center gap-2 mt-0.5">
                      <span className="text-[11px] text-neutral-500">{formatDate(item.date)}</span>
                      <Badge className={item.cancelled ? 'bg-red-50 text-red-600 text-[10px]' : 'bg-neutral-100 text-neutral-600 text-[10px]'}>
                        {item.cancelled ? '취소됨' : item.sub}
                      </Badge>
                    </div>
                  </div>
                </div>
              );
            })
          )}
        </div>
      )}
    </div>
  );
}
