'use client';

import { use, useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomer, useUpdateCustomer } from '@/hooks/use-customers';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { ArrowLeft, Save, ShoppingBag, FileSignature, MessageSquare, Wrench, Clock } from 'lucide-react';

const TYPE_OPTIONS = [
  { value: 'retail', label: '일반' },
  { value: 'online', label: '온라인' },
  { value: 'dealer', label: '딜러' },
  { value: 'academy', label: '아카데미' },
];

const PAYMENT_STATUS_LABEL: Record<string, string> = {
  paid: '결제완료',
  unpaid: '미결제',
  partial: '부분결제',
};

const CONSULTATION_TYPE_LABEL: Record<string, string> = {
  store_visit: '매장방문',
  field_request: '출장',
  talk_consult: '톡상담',
};

const CONTRACT_STATUS_LABEL: Record<string, string> = {
  draft: '작성중',
  signed: '서명완료',
  sent: '발송',
  completed: '완료',
  cancelled: '취소',
};

export default function CustomerDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();
  const { data, isLoading } = useCustomer(id);
  const updateCustomer = useUpdateCustomer();

  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState({
    name: '',
    phone: '',
    email: '',
    customer_type: 'retail',
    company_name: '',
    address_road: '',
    address_detail: '',
    memo: '',
    outstanding_balance: 0,
  });

  function startEdit() {
    if (!data?.customer) return;
    const c = data.customer;
    setForm({
      name: c.name,
      phone: c.phone || '',
      email: c.email || '',
      customer_type: c.customer_type || 'retail',
      company_name: c.company_name || '',
      address_road: c.address_road || '',
      address_detail: c.address_detail || '',
      memo: c.memo || '',
      outstanding_balance: c.outstanding_balance || 0,
    });
    setEditing(true);
  }

  async function handleSave() {
    await updateCustomer.mutateAsync({ id, ...form });
    setEditing(false);
  }

  if (isLoading) {
    return (
      <>
        <Topbar title="고객 상세" />
        <div className="px-4 md:px-6 py-4 space-y-4">
          <Skeleton className="h-48" />
          <Skeleton className="h-32" />
        </div>
      </>
    );
  }

  if (!data?.customer) {
    return (
      <>
        <Topbar title="고객 상세" />
        <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
          고객 정보를 찾을 수 없습니다
        </div>
      </>
    );
  }

  const { customer: c, sales, contracts, consultations, repairs, summary } = data as Record<string, unknown> & typeof data;

  // 통합 타임라인 생성
  type TimelineItem = { type: 'sale' | 'contract' | 'consultation' | 'repair'; date: string; id: string; title: string; subtitle: string; amount?: number; status?: string; href?: string };
  const timeline: TimelineItem[] = [
    ...sales.map((s) => ({ type: 'sale' as const, date: s.sale_date, id: s.id, title: s.sale_number, subtitle: `판매 · ${PAYMENT_STATUS_LABEL[s.payment_status] || s.payment_status}`, amount: s.paid_amount, href: `/sales/${s.id}` })),
    ...contracts.map((ct) => ({ type: 'contract' as const, date: ct.created_at, id: ct.id, title: ct.contract_number, subtitle: `계약서 · ${CONTRACT_STATUS_LABEL[ct.status] || ct.status}`, amount: ct.final_amount, href: `/contracts/${ct.id}` })),
    ...consultations.map((con) => ({ type: 'consultation' as const, date: con.visit_date || con.created_at, id: con.id, title: CONSULTATION_TYPE_LABEL[con.consultation_type] || con.consultation_type, subtitle: `상담 · ${con.status}`, href: `/consultations/${con.id}` })),
    ...((repairs as unknown as Array<{ id: string; repair_number: string; status: string; total_cost: number | null; created_at: string }>) || []).map((r) => ({ type: 'repair' as const, date: r.created_at, id: r.id, title: r.repair_number || '복원수리', subtitle: `복원수리 · ${r.status}`, amount: r.total_cost || 0, href: `/repairs/${r.id}` })),
  ].sort((a, b) => b.date.localeCompare(a.date));

  const TIMELINE_ICON = { sale: ShoppingBag, contract: FileSignature, consultation: MessageSquare, repair: Wrench };
  const TIMELINE_COLOR = { sale: 'text-green-600 bg-green-50', contract: 'text-purple-600 bg-purple-50', consultation: 'text-blue-600 bg-blue-50', repair: 'text-orange-600 bg-orange-50' };

  // RFM 분류
  const lastSaleDate = (summary as Record<string, unknown>).lastSaleDate as string | null;
  const daysSinceLastSale = lastSaleDate ? Math.floor((Date.now() - new Date(lastSaleDate).getTime()) / (1000 * 60 * 60 * 24)) : 999;
  const rfmLabel = summary.totalSales >= 3 || summary.totalSalesAmount >= 500000 ? 'VIP'
    : daysSinceLastSale > 180 ? '휴면'
    : '일반';
  const rfmColor = rfmLabel === 'VIP' ? 'bg-amber-100 text-amber-700' : rfmLabel === '휴면' ? 'bg-neutral-200 text-neutral-500' : 'bg-blue-100 text-blue-700';

  return (
    <>
      <Topbar title="고객 상세" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <Button variant="ghost" size="sm" onClick={() => router.push('/customers')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>

        {/* 고객 정보 카드 */}
        <Card>
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-bold text-indigo-black">{c.name}</h3>
            {!editing ? (
              <Button variant="ghost" size="sm" onClick={startEdit}>수정</Button>
            ) : (
              <Button size="sm" onClick={handleSave} disabled={updateCustomer.isPending}>
                <Save size={14} />
                {updateCustomer.isPending ? '저장 중...' : '저장'}
              </Button>
            )}
          </div>

          {editing ? (
            <div className="space-y-3">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-neutral-500">고객명</label>
                  <input
                    type="text"
                    value={form.name}
                    onChange={(e) => setForm({ ...form, name: e.target.value })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">연락처</label>
                  <input
                    type="text"
                    value={form.phone}
                    onChange={(e) => setForm({ ...form, phone: e.target.value })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">이메일</label>
                  <input
                    type="email"
                    value={form.email}
                    onChange={(e) => setForm({ ...form, email: e.target.value })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">고객 유형</label>
                  <select
                    value={form.customer_type}
                    onChange={(e) => setForm({ ...form, customer_type: e.target.value })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  >
                    {TYPE_OPTIONS.map((t) => (
                      <option key={t.value} value={t.value}>{t.label}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div>
                <label className="text-xs text-neutral-500">매장명 (근무지)</label>
                <input
                  type="text"
                  value={form.company_name}
                  onChange={(e) => setForm({ ...form, company_name: e.target.value })}
                  placeholder="매장 또는 근무지명"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                />
              </div>
              <div>
                <label className="text-xs text-neutral-500">주소</label>
                <input
                  type="text"
                  value={form.address_road}
                  onChange={(e) => setForm({ ...form, address_road: e.target.value })}
                  placeholder="도로명 주소"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                />
                <input
                  type="text"
                  value={form.address_detail}
                  onChange={(e) => setForm({ ...form, address_detail: e.target.value })}
                  placeholder="상세 주소 (동/호수)"
                  className="w-full h-9 px-3 mt-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                />
              </div>
              <div>
                <label className="text-xs text-neutral-500">메모</label>
                <textarea
                  value={form.memo}
                  onChange={(e) => setForm({ ...form, memo: e.target.value })}
                  rows={2}
                  className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none"
                />
              </div>
              <div>
                <label className="text-xs text-neutral-500">미수금</label>
                <input
                  type="number"
                  value={form.outstanding_balance || ''}
                  onChange={(e) => setForm({ ...form, outstanding_balance: parseInt(e.target.value) || 0 })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                />
              </div>
              <Button variant="ghost" size="sm" onClick={() => setEditing(false)}>취소</Button>
            </div>
          ) : (
            <>
              <div className="grid grid-cols-2 gap-3 text-sm">
                <div>
                  <span className="text-xs text-neutral-500">연락처</span>
                  <p>{formatPhone(c.phone) || '-'}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">이메일</span>
                  <p>{c.email || '-'}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">유형</span>
                  <p>
                    <Badge className={
                      c.customer_type === 'dealer' ? 'bg-purple-100 text-purple-700'
                        : c.customer_type === 'academy' ? 'bg-emerald-100 text-emerald-700'
                        : c.customer_type === 'online' ? 'bg-blue-100 text-blue-700'
                        : 'bg-neutral-100 text-neutral-600'
                    }>
                      {TYPE_OPTIONS.find((t) => t.value === c.customer_type)?.label || '일반'}
                    </Badge>
                  </p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">등록일</span>
                  <p>{formatDate(c.created_at)}</p>
                </div>
                {c.company_name && (
                  <div className="col-span-2">
                    <span className="text-xs text-neutral-500">매장명 (근무지)</span>
                    <p>{c.company_name}</p>
                  </div>
                )}
                {(c.address_road || c.address_detail) && (
                  <div className="col-span-2">
                    <span className="text-xs text-neutral-500">주소</span>
                    <p>{[c.address_road, c.address_detail].filter(Boolean).join(' ')}</p>
                  </div>
                )}
              </div>
              {c.memo && (
                <p className="mt-3 pt-3 border-t border-neutral-100 text-sm text-neutral-600">{c.memo}</p>
              )}
            </>
          )}
        </Card>

        {/* RFM 뱃지 + 요약 카드 */}
        <div className="flex items-center gap-2 mb-1">
          <span className={`px-2 py-0.5 rounded text-xs font-bold ${rfmColor}`}>{rfmLabel}</span>
          {daysSinceLastSale < 999 && <span className="text-xs text-neutral-400">마지막 거래 {daysSinceLastSale}일 전</span>}
        </div>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
          <Card>
            <p className="text-xs text-neutral-500">총 판매건수</p>
            <p className="text-lg font-bold text-indigo-black">{summary.totalSales}</p>
          </Card>
          <Card>
            <p className="text-xs text-neutral-500">총 판매액</p>
            <p className="text-lg font-bold text-terracotta">{formatKRW(summary.totalSalesAmount)}</p>
          </Card>
          <Card>
            <p className="text-xs text-neutral-500">미수금</p>
            <p className={`text-lg font-bold ${c.outstanding_balance > 0 ? 'text-red-500' : 'text-neutral-400'}`}>
              {formatKRW(c.outstanding_balance)}
            </p>
          </Card>
          <Card>
            <p className="text-xs text-neutral-500">최근 거래일</p>
            <p className="text-sm font-bold text-neutral-700">
              {(summary as Record<string, unknown>).lastSaleDate ? formatDate((summary as Record<string, unknown>).lastSaleDate as string) : '-'}
            </p>
          </Card>
        </div>

        {/* 통합 타임라인 */}
        {timeline.length > 0 && (
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3 flex items-center gap-2">
              <Clock size={16} />
              거래 이력 ({timeline.length})
            </h3>
            <div className="space-y-1">
              {timeline.map((item) => {
                const Icon = TIMELINE_ICON[item.type];
                const color = TIMELINE_COLOR[item.type];
                return (
                  <div
                    key={`${item.type}-${item.id}`}
                    onClick={() => item.href && router.push(item.href)}
                    className={`flex items-center gap-3 py-2.5 border-b border-neutral-50 last:border-0 ${item.href ? 'cursor-pointer hover:bg-warm-ivory/40' : ''} -mx-1 px-1 rounded transition`}
                  >
                    <div className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 ${color}`}>
                      <Icon size={14} />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium truncate">{item.title}</p>
                      <p className="text-xs text-neutral-500">{item.subtitle} · {formatDate(item.date)}</p>
                    </div>
                    {item.amount !== undefined && item.amount > 0 && (
                      <p className="text-sm font-bold text-neutral-800 shrink-0">{formatKRW(item.amount)}</p>
                    )}
                  </div>
                );
              })}
            </div>
          </Card>
        )}
      </div>
    </>
  );
}
