'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomer, useUpdateCustomer } from '@/hooks/use-customers';
import { useCustomerManualInvoices } from '@/hooks/use-manual-invoices';
import { formatKRW, formatDate, formatPhone, ORDER_STATUS_LABEL, CUSTOMER_TYPE_COLOR, CONSULTATION_TYPE_LABEL } from '@/lib/utils/format';
import { activitySuffix } from '@/lib/customer/display';
import {
  Save, ShoppingBag, FileSignature, MessageSquare, Wrench,
  Clock, Pencil, X, Truck, Copy, Merge, Package,
} from 'lucide-react';
import { TagBadges, TagSelector } from '@/components/shared/tag-selector';
import { CustomerCreateModal } from '@/components/customers/customer-create-modal';
import { CustomerNotes } from '@/components/shared/customer-notes';
import { CustomerMergeModal } from '@/components/customers/customer-merge-modal';
import { DaumPostcodeButton } from '@/components/shared/daum-postcode-button';
import { useSetting } from '@/hooks/use-settings';

const TYPE_OPTIONS = [
  { value: 'retail', label: '일반' },
  { value: 'online', label: '온라인' },
  { value: 'dealer', label: '딜러' },
  { value: 'academy', label: '아카데미' },
];

// 고객유형 색·상담유형 라벨은 format.ts SSOT(CUSTOMER_TYPE_COLOR/CONSULTATION_TYPE_LABEL) 사용
const PAYMENT_STATUS_LABEL: Record<string, string> = {
  paid: '결제완료', unpaid: '미결제', partial: '부분결제',
};

const REPAIR_STATUS_LABEL: Record<string, string> = {
  intake: '접수', pickup_scheduled: '수거접수', cost_notified: '비용안내',
  repairing: '작업중', shipped: '출고', delivered: '배송완료',
  completed: '완료', cancelled: '취소',
};

interface Props {
  customerId: string;
  /** 풀페이지(/customers/[id])에서 래퍼로 쓸 때 '상세 페이지' 자기링크 버튼 숨김 */
  hideDetailLink?: boolean;
}

const EMPTY_FORM = {
  name: '', phone: '', email: '', customer_type: 'retail',
  activity_name: '', position: '', company_name: '',
  postcode: '', address_road: '', address_detail: '',
  outstanding_balance: 0, default_repair_price: 0,
  tags: [] as string[],
};

export function CustomerDetailPanel({ customerId, hideDetailLink }: Props) {
  const router = useRouter();
  const { data, isLoading } = useCustomer(customerId);
  const { data: manualInvoicesData } = useCustomerManualInvoices(customerId);
  const updateCustomer = useUpdateCustomer();
  const availableTags = useSetting<string[]>('customer.tags', []);
  const [tab, setTab] = useState<'profile' | 'timeline'>('profile');
  const [editingMemo, setEditingMemo] = useState(false);
  const [editingInfo, setEditingInfo] = useState(false);
  const [form, setForm] = useState(EMPTY_FORM);
  const [showDuplicate, setShowDuplicate] = useState(false);
  const [showMerge, setShowMerge] = useState(false);
  const [editMemo, setEditMemo] = useState('');

  if (isLoading) {
    return <div className="space-y-4"><Skeleton className="h-32" /><Skeleton className="h-24" /><Skeleton className="h-48" /></div>;
  }
  if (!data?.customer) {
    return <p className="text-sm text-neutral-400 text-center py-8">고객 정보를 찾을 수 없습니다</p>;
  }

  const { customer: c, sales, contracts, consultations, repairs, summary } = data;
  const orders = data.orders ?? [];
  const manualInvoices = manualInvoicesData?.invoices ?? [];

  // 통합 타임라인 생성 — 시간순 역순
  const timeline = [
    ...orders.map((ord) => ({
      type: 'order' as const,
      id: ord.id,
      date: ord.ordered_at,
      title: ord.imweb_order_no,
      sub: ORDER_STATUS_LABEL[ord.status] || ord.status,
      amount: ord.paid_amount,
      icon: Package,
      color: 'text-emerald-600 bg-emerald-50',
      cancelled: ord.status === 'cancelled',
    })),
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
    setEditingMemo(true);
  }

  async function saveMemo() {
    await updateCustomer.mutateAsync({ id: customerId, memo: editMemo });
    setEditingMemo(false);
  }

  function startInfoEdit() {
    setForm({
      name: c.name,
      phone: c.phone || '',
      email: c.email || '',
      customer_type: c.customer_type || 'retail',
      activity_name: c.activity_name || '',
      position: c.position || '',
      company_name: c.company_name || '',
      postcode: c.postcode || '',
      address_road: c.address_road || '',
      address_detail: c.address_detail || '',
      outstanding_balance: c.outstanding_balance || 0,
      default_repair_price: (c as Record<string, unknown>).default_repair_price as number || 0,
      tags: (c as unknown as { tags?: string[] }).tags || [],
    });
    setEditingInfo(true);
  }
  async function saveInfo() {
    await updateCustomer.mutateAsync({
      id: customerId,
      ...form,
      activity_name: form.activity_name.trim() || null,
      position: form.position.trim() || null,
    });
    setEditingInfo(false);
  }

  // RFM 분류 (풀페이지에서 흡수)
  const lastSaleDate = (summary as Record<string, unknown>).lastSaleDate as string | null;
  const daysSinceLastSale = lastSaleDate ? Math.floor((Date.now() - new Date(lastSaleDate).getTime()) / (1000 * 60 * 60 * 24)) : 999;
  const rfmLabel = summary.totalSales >= 3 || summary.totalSalesAmount >= 500000 ? 'VIP'
    : daysSinceLastSale > 180 ? '휴면' : '일반';
  const rfmColor = rfmLabel === 'VIP' ? 'bg-amber-100 text-amber-700' : rfmLabel === '휴면' ? 'bg-neutral-200 text-neutral-500' : 'bg-blue-100 text-blue-700';

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
            <h2 className="text-base font-bold text-stone-900">
              {c.name}
              {activitySuffix(c.activity_name, c.position) && (
                <span className="ml-1.5 text-xs font-normal text-neutral-400">{activitySuffix(c.activity_name, c.position)}</span>
              )}
            </h2>
            <Badge className={CUSTOMER_TYPE_COLOR[c.customer_type] || CUSTOMER_TYPE_COLOR.retail}>
              {TYPE_OPTIONS.find(t => t.value === c.customer_type)?.label || '일반'}
            </Badge>
            <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${rfmColor}`}>{rfmLabel}</span>
          </div>
          <p className="text-xs text-neutral-500 mt-0.5">
            {formatPhone(c.phone) || '연락처 없음'}
            {c.company_name && ` · ${c.company_name}`}
            {daysSinceLastSale < 999 && <span className="text-neutral-400"> · 마지막 거래 {daysSinceLastSale}일 전</span>}
          </p>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          <Button variant="secondary" size="sm" onClick={() => setShowDuplicate(true)} title="같은 매장 동료 등록 (매장명·주소 자동)">
            <Copy size={13} />
            복제
          </Button>
          <Button variant="secondary" size="sm" onClick={() => setShowMerge(true)} title="중복 고객을 이 고객으로 병합">
            <Merge size={13} />
            병합
          </Button>
          {!hideDetailLink && (
            <Button variant="ghost" size="sm" onClick={() => router.push(`/customers/${customerId}`)}>
              상세 페이지
            </Button>
          )}
        </div>
      </div>

      {/* 복제 등록 — 같은 매장(매장명·주소·유형·태그) 채워서 새 고객 등록 */}
      <CustomerCreateModal
        open={showDuplicate}
        onClose={() => setShowDuplicate(false)}
        onCreated={() => setShowDuplicate(false)}
        prefill={{
          customer_type: c.customer_type,
          company_name: c.company_name || '',
          postcode: c.postcode || '',
          address_road: c.address_road || '',
          address_detail: c.address_detail || '',
          tags: (c as unknown as { tags?: string[] }).tags || [],
        }}
        // 복제는 같은 매장 동료 — 활동명/직급은 사람마다 달라 비움(매장명·주소만 채움)
      />

      {/* 병합 — 중복 고객을 이 고객으로 흡수 */}
      <CustomerMergeModal
        open={showMerge}
        onClose={() => setShowMerge(false)}
        primary={{ id: customerId, name: c.name, phone: c.phone }}
      />

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
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-xs font-bold text-stone-900">기본 정보</h3>
              {!editingInfo ? (
                <button onClick={startInfoEdit} className="text-neutral-400 hover:text-neutral-600" title="고객 정보 수정"><Pencil size={14} /></button>
              ) : (
                <div className="flex gap-1">
                  <button onClick={() => setEditingInfo(false)} className="text-neutral-400 hover:text-neutral-600"><X size={14} /></button>
                  <button onClick={saveInfo} disabled={updateCustomer.isPending} className="text-blue-600 hover:text-blue-700"><Save size={14} /></button>
                </div>
              )}
            </div>

            {editingInfo ? (
              <div className="space-y-3">
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="text-xs text-neutral-500">고객명</label>
                    <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  </div>
                  <div>
                    <label className="text-xs text-neutral-500">연락처</label>
                    <input type="text" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  </div>
                  <div>
                    <label className="text-xs text-neutral-500">이메일</label>
                    <input type="email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  </div>
                  <div>
                    <label className="text-xs text-neutral-500">고객 유형</label>
                    <select value={form.customer_type} onChange={(e) => setForm({ ...form, customer_type: e.target.value })}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400">
                      {TYPE_OPTIONS.map((t) => <option key={t.value} value={t.value}>{t.label}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-xs text-neutral-500">활동명 (매장 사용 이름)</label>
                    <input type="text" value={form.activity_name} onChange={(e) => setForm({ ...form, activity_name: e.target.value })} placeholder="예) 하은"
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  </div>
                  <div>
                    <label className="text-xs text-neutral-500">직급</label>
                    <input type="text" value={form.position} onChange={(e) => setForm({ ...form, position: e.target.value })} placeholder="예) 디자이너"
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  </div>
                </div>
                <div>
                  <label className="text-xs text-neutral-500">매장명 (근무지)</label>
                  <input type="text" value={form.company_name} onChange={(e) => setForm({ ...form, company_name: e.target.value })} placeholder="매장 또는 근무지명"
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">주소</label>
                  <div className="flex gap-2 mb-2">
                    <input type="text" value={form.postcode} readOnly placeholder="우편번호"
                      className="w-24 h-9 px-3 rounded-lg border border-neutral-200 bg-neutral-50 text-sm text-neutral-600" />
                    <DaumPostcodeButton onSelected={(d) => setForm((prev) => ({ ...prev, postcode: d.zonecode, address_road: d.roadAddress }))}>
                      주소검색
                    </DaumPostcodeButton>
                  </div>
                  <input type="text" value={form.address_road} readOnly placeholder="도로명 주소"
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-neutral-50 text-sm text-neutral-600" />
                  <input type="text" value={form.address_detail} onChange={(e) => setForm({ ...form, address_detail: e.target.value })} placeholder="상세 주소 (동/호수)"
                    className="w-full h-9 px-3 mt-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
                </div>
                {availableTags.length > 0 && (
                  <div>
                    <label className="text-xs text-neutral-500 mb-1 block">태그</label>
                    <TagSelector availableTags={availableTags} selectedTags={form.tags} onChange={(tags) => setForm({ ...form, tags })} />
                  </div>
                )}
                <div>
                  <label className="text-xs text-neutral-500">미수금</label>
                  <input type="number" value={form.outstanding_balance || ''} onChange={(e) => setForm({ ...form, outstanding_balance: parseInt(e.target.value) || 0 })}
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                </div>
                {(form.customer_type === 'dealer' || form.customer_type === 'academy') && (
                  <div>
                    <label className="text-xs text-neutral-500">복원수리 기본 단가 (자루당)</label>
                    <input type="number" value={form.default_repair_price || ''} onChange={(e) => setForm({ ...form, default_repair_price: parseInt(e.target.value) || 0 })}
                      placeholder="예: 8000 — 납품 복원수리 입력 시 자동 적용"
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
                    <p className="text-[10px] text-neutral-400 mt-1">납품 → 복원수리 입력 시 이 단가가 자동으로 채워집니다 (비워두면 기본 8,000원)</p>
                  </div>
                )}
              </div>
            ) : (
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
                  <span className="text-xs text-neutral-500">활동명 (매장 사용 이름)</span>
                  <p>{c.activity_name || '-'}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">직급</span>
                  <p>{c.position || '-'}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">등록일</span>
                  <p>{formatDate(c.created_at)}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">매장명 (근무지)</span>
                  <p>{c.company_name || '-'}</p>
                </div>
                {(c.customer_type === 'dealer' || c.customer_type === 'academy') && (
                  <div>
                    <span className="text-xs text-neutral-500">복원수리 기본 단가</span>
                    <p>{(c as Record<string, unknown>).default_repair_price ? `${formatKRW((c as Record<string, unknown>).default_repair_price as number)} / 자루` : <span className="text-neutral-400">미설정 (기본 8,000원)</span>}</p>
                  </div>
                )}
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
            )}
          </Card>

          {/* 메모 */}
          <Card>
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-xs font-bold text-stone-900">메모</h3>
              {!editingMemo ? (
                <button onClick={startMemoEdit} className="text-neutral-400 hover:text-neutral-600">
                  <Pencil size={14} />
                </button>
              ) : (
                <div className="flex gap-1">
                  <button onClick={() => setEditingMemo(false)} className="text-neutral-400 hover:text-neutral-600">
                    <X size={14} />
                  </button>
                  <button onClick={saveMemo} disabled={updateCustomer.isPending}
                    className="text-blue-600 hover:text-blue-700">
                    <Save size={14} />
                  </button>
                </div>
              )}
            </div>
            {editingMemo ? (
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

          {/* 상담 메모 타임라인 (특징·불편·요구 등 날짜별 기록) */}
          <Card>
            <h3 className="text-xs font-bold text-stone-900 mb-2">상담 메모</h3>
            <CustomerNotes customerId={customerId} />
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
                    else if (item.type === 'order') router.push(`/orders/${item.id}`);
                    else if (item.type === 'manual_invoice') router.push('/manual-invoices');
                  }}
                  className={`flex items-start gap-3 py-3 ${
                    i < timeline.length - 1 ? 'border-b border-neutral-100' : ''
                  } ${item.type === 'sale' || item.type === 'contract' || item.type === 'order' || item.type === 'manual_invoice' ? 'cursor-pointer hover:bg-stone-50/40 -mx-1 px-1 rounded' : ''}`}
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
