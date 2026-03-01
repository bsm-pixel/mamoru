'use client';

import { use, useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomer, useUpdateCustomer } from '@/hooks/use-customers';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { ArrowLeft, Save, ShoppingBag, FileSignature, MessageSquare } from 'lucide-react';

const TYPE_OPTIONS = [
  { value: 'retail', label: '일반' },
  { value: 'online', label: '온라인' },
  { value: 'dealer', label: '딜러' },
  { value: 'supplier', label: '매입처' },
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

  const { customer: c, sales, contracts, consultations, summary } = data;

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
                <label className="text-xs text-neutral-500">업체명</label>
                <input
                  type="text"
                  value={form.company_name}
                  onChange={(e) => setForm({ ...form, company_name: e.target.value })}
                  placeholder="해당 시 입력"
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
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
                  <p>{c.phone || '-'}</p>
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
                        : c.customer_type === 'supplier' ? 'bg-amber-100 text-amber-700'
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
                    <span className="text-xs text-neutral-500">업체명</span>
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

        {/* 요약 카드 */}
        <div className="grid grid-cols-3 gap-3">
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
        </div>

        {/* 판매내역 */}
        {sales.length > 0 && (
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3 flex items-center gap-2">
              <ShoppingBag size={16} />
              판매내역 ({sales.length})
            </h3>
            <div className="space-y-2">
              {sales.map((s) => (
                <div
                  key={s.id}
                  onClick={() => router.push(`/sales/${s.id}`)}
                  className="flex items-center justify-between py-2 border-b border-neutral-50 last:border-0 cursor-pointer hover:bg-warm-ivory/40 -mx-1 px-1 rounded transition"
                >
                  <div>
                    <p className="text-sm font-medium">{s.sale_number}</p>
                    <p className="text-xs text-neutral-500">{formatDate(s.sale_date)}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-sm font-bold">{formatKRW(s.paid_amount)}</p>
                    <p className="text-xs text-neutral-500">{PAYMENT_STATUS_LABEL[s.payment_status] || s.payment_status}</p>
                  </div>
                </div>
              ))}
            </div>
          </Card>
        )}

        {/* 계약서 */}
        {contracts.length > 0 && (
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3 flex items-center gap-2">
              <FileSignature size={16} />
              계약서 ({contracts.length})
            </h3>
            <div className="space-y-2">
              {contracts.map((ct) => (
                <div
                  key={ct.id}
                  onClick={() => router.push(`/contracts/${ct.id}`)}
                  className="flex items-center justify-between py-2 border-b border-neutral-50 last:border-0 cursor-pointer hover:bg-warm-ivory/40 -mx-1 px-1 rounded transition"
                >
                  <div>
                    <p className="text-sm font-medium">{ct.contract_number}</p>
                    <p className="text-xs text-neutral-500">{formatDate(ct.created_at)}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-sm font-bold">{formatKRW(ct.final_amount)}</p>
                    <p className="text-xs text-neutral-500">{CONTRACT_STATUS_LABEL[ct.status] || ct.status}</p>
                  </div>
                </div>
              ))}
            </div>
          </Card>
        )}

        {/* 상담내역 */}
        {consultations.length > 0 && (
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3 flex items-center gap-2">
              <MessageSquare size={16} />
              상담내역 ({consultations.length})
            </h3>
            <div className="space-y-2">
              {consultations.map((con) => (
                <div
                  key={con.id}
                  className="flex items-center justify-between py-2 border-b border-neutral-50 last:border-0"
                >
                  <div>
                    <p className="text-sm font-medium">
                      {CONSULTATION_TYPE_LABEL[con.consultation_type] || con.consultation_type}
                    </p>
                    <p className="text-xs text-neutral-500">
                      {con.visit_date ? formatDate(con.visit_date) : formatDate(con.created_at)}
                    </p>
                  </div>
                  <Badge className="bg-neutral-100 text-neutral-600">{con.status}</Badge>
                </div>
              ))}
            </div>
          </Card>
        )}
      </div>
    </>
  );
}
