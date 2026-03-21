'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomers, useCreateCustomer } from '@/hooks/use-customers';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { Users, Plus, X } from 'lucide-react';
import type { Customer } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

const TYPE_LABEL: Record<string, string> = {
  retail: '일반',
  online: '온라인',
  dealer: '딜러',
  academy: '아카데미',
};

const TYPE_COLOR: Record<string, string> = {
  retail: 'bg-neutral-100 text-neutral-600',
  online: 'bg-blue-100 text-blue-700',
  dealer: 'bg-purple-100 text-purple-700',
  academy: 'bg-emerald-100 text-emerald-700',
};

const SOURCE_LABEL: Record<string, string> = {
  imweb: '아임웹',
  consultation: '상담',
  as: '복원수리',
  manual: '수동',
};

const FILTER_TYPES = [
  { value: '', label: '전체' },
  { value: 'retail', label: '일반' },
  { value: 'online', label: '온라인' },
  { value: 'dealer', label: '딜러' },
  { value: 'academy', label: '아카데미' },
];

export default function CustomersPage() {
  const router = useRouter();
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('');
  const [page, setPage] = useState(1);
  const [showAdd, setShowAdd] = useState(false);
  const limit = 20;

  const { data, isLoading } = useCustomers({
    search,
    type: typeFilter || undefined,
    page,
    limit,
  });
  const customers = data?.customers || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);

  return (
    <>
      <Topbar title="고객 관리" action={
        <Button size="sm" onClick={() => setShowAdd(true)}>
          <Plus size={14} />
          고객 추가
        </Button>
      } />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 검색 */}
        <SearchInput
          value={search}
          onChange={(v) => { setSearch(v); setPage(1); }}
          placeholder="고객명, 전화번호, 업체명 검색"
        />

        {/* 유형 필터 */}
        <div className="flex gap-1.5 overflow-x-auto pb-1">
          {FILTER_TYPES.map((f) => (
            <button
              key={f.value}
              onClick={() => { setTypeFilter(f.value); setPage(1); }}
              className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                typeFilter === f.value
                  ? 'bg-terracotta text-cream'
                  : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>

        {/* 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-16 w-full" />
              ))}
            </div>
          ) : customers.length === 0 ? (
            <EmptyState icon={Users} message="고객이 없습니다" />
          ) : (
            <div className="divide-y divide-neutral-100">
              {customers.map((c) => (
                <CustomerRow
                  key={c.id}
                  customer={c}
                  onClick={() => router.push(`/customers/${c.id}`)}
                />
              ))}
            </div>
          )}
        </Card>

        <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} unit="명" />
      </div>

      {showAdd && <AddCustomerModal onClose={() => setShowAdd(false)} />}
    </>
  );
}

function AddCustomerModal({ onClose }: { onClose: () => void }) {
  const createCustomer = useCreateCustomer();
  const [form, setForm] = useState({
    name: '', phone: '', customerType: 'retail',
    companyName: '', address: '', addressDetail: '', memo: '',
  });

  const inputCls = "w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40";

  async function handleSubmit() {
    if (!form.name.trim()) { toast.error('고객명을 입력해주세요'); return; }
    await createCustomer.mutateAsync({
      name: form.name.trim(),
      phone: form.phone.trim() || undefined,
      customerType: form.customerType,
      companyName: form.companyName.trim() || undefined,
      address: form.address.trim() || undefined,
      addressDetail: form.addressDetail.trim() || undefined,
      memo: form.memo.trim() || undefined,
    });
    onClose();
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl w-full max-w-md mx-4 shadow-xl max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-100 sticky top-0 bg-white rounded-t-xl z-10">
          <h2 className="text-sm font-bold text-indigo-black">고객 추가</h2>
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center"><X size={16} /></button>
        </div>
        <div className="px-5 py-4 space-y-4">
          {/* 기본 정보 */}
          <div className="grid grid-cols-2 gap-3">
            <div className="col-span-2 sm:col-span-1">
              <label className="text-xs text-neutral-500">고객명 *</label>
              <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="고객 이름" autoFocus className={inputCls} />
            </div>
            <div className="col-span-2 sm:col-span-1">
              <label className="text-xs text-neutral-500">고객 유형</label>
              <select value={form.customerType} onChange={(e) => setForm({ ...form, customerType: e.target.value })}
                className={inputCls}>
                {FILTER_TYPES.filter(f => f.value).map(f => (
                  <option key={f.value} value={f.value}>{f.label}</option>
                ))}
              </select>
            </div>
          </div>

          <div>
            <label className="text-xs text-neutral-500">전화번호</label>
            <input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
              placeholder="010-0000-0000" className={inputCls} />
          </div>

          <div>
            <label className="text-xs text-neutral-500">매장명 (근무지)</label>
            <input type="text" value={form.companyName} onChange={(e) => setForm({ ...form, companyName: e.target.value })}
              placeholder="매장 또는 근무지명" className={inputCls} />
          </div>

          {/* 주소 */}
          <div className="space-y-2">
            <label className="text-xs text-neutral-500">주소</label>
            <input type="text" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })}
              placeholder="도로명 주소" className={inputCls} />
            <input type="text" value={form.addressDetail} onChange={(e) => setForm({ ...form, addressDetail: e.target.value })}
              placeholder="상세 주소 (동/호수)" className={inputCls} />
          </div>

          {/* 메모 */}
          <div>
            <label className="text-xs text-neutral-500">메모</label>
            <textarea value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })}
              placeholder="특이사항, 참고사항 등" rows={2}
              className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none" />
          </div>
        </div>
        <div className="px-5 py-4 border-t border-neutral-100 flex gap-2 sticky bottom-0 bg-white rounded-b-xl">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button className="flex-1" disabled={!form.name.trim() || createCustomer.isPending} onClick={handleSubmit}>
            {createCustomer.isPending ? '등록 중...' : '고객 등록'}
          </Button>
        </div>
      </div>
    </div>
  );
}

function CustomerRow({ customer, onClick }: { customer: Customer; onClick: () => void }) {
  const c = customer;
  return (
    <div
      onClick={onClick}
      className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition"
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
          <Badge className={TYPE_COLOR[c.customer_type] || TYPE_COLOR.retail}>
            {TYPE_LABEL[c.customer_type] || '일반'}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          {c.phone && <span>{c.phone}</span>}
          {c.company_name && <span>{c.company_name}</span>}
          <span>{SOURCE_LABEL[c.source] || c.source}</span>
          <span>{formatDate(c.created_at)}</span>
        </div>
      </div>
      <div className="text-right shrink-0">
        {c.total_spent > 0 && (
          <p className="text-sm font-bold">{formatKRW(c.total_spent)}</p>
        )}
        {c.outstanding_balance > 0 && (
          <p className="text-xs text-red-500 font-semibold">미수 {formatKRW(c.outstanding_balance)}</p>
        )}
      </div>
    </div>
  );
}
