'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomers, useCreateCustomer } from '@/hooks/use-customers';
import { formatKRW } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Building2, Plus, X } from 'lucide-react';
import type { Customer } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

export default function SuppliersPage() {
  const [search, setSearch] = useState('');
  const [showAdd, setShowAdd] = useState(false);

  const { data, isLoading } = useCustomers({ type: 'supplier', search, limit: 100 });
  const suppliers = data?.customers || [];

  return (
    <>
      <Topbar title="매입처 관리" action={
        <Button size="sm" onClick={() => setShowAdd(true)}>
          <Plus size={14} />
          매입처 추가
        </Button>
      } />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <SearchInput
          value={search}
          onChange={setSearch}
          placeholder="매입처명, 담당자명, 전화번호 검색"
        />

        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 4 }).map((_, i) => <Skeleton key={i} className="h-16 w-full" />)}
            </div>
          ) : suppliers.length === 0 ? (
            <EmptyState icon={Building2} message="등록된 매입처가 없습니다" />
          ) : (
            <div className="divide-y divide-neutral-100">
              {suppliers.map((s) => (
                <SupplierRow key={s.id} supplier={s} />
              ))}
            </div>
          )}
        </Card>

        <p className="text-xs text-neutral-400">총 {suppliers.length}개 매입처</p>
      </div>

      {showAdd && <AddSupplierModal onClose={() => setShowAdd(false)} />}
    </>
  );
}

function SupplierRow({ supplier: s }: { supplier: Customer }) {
  return (
    <div className="flex items-center gap-4 px-4 py-3 hover:bg-warm-ivory/60 transition">
      <div className="w-9 h-9 rounded-lg bg-amber-50 flex items-center justify-center shrink-0">
        <Building2 size={18} className="text-amber-600" />
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black">{s.company_name || s.name}</span>
          <Badge className="bg-amber-100 text-amber-700">매입처</Badge>
        </div>
        <div className="flex items-center gap-3 mt-0.5 text-xs text-neutral-500">
          {s.company_name && <span>담당: {s.name}</span>}
          {s.phone && <span>{s.phone}</span>}
          {s.memo && <span className="truncate max-w-[200px]">{s.memo}</span>}
        </div>
      </div>
      <div className="text-right shrink-0">
        {s.total_spent > 0 && (
          <p className="text-sm font-bold">{formatKRW(s.total_spent)}</p>
        )}
      </div>
    </div>
  );
}

function AddSupplierModal({ onClose }: { onClose: () => void }) {
  const createCustomer = useCreateCustomer();
  const [form, setForm] = useState({
    companyName: '', name: '', phone: '', memo: '',
    businessNumber: '', representative: '', businessType: '', businessCategory: '',
    email: '', address: '', contactChannel: '',
  });

  async function handleSubmit() {
    if (!form.companyName.trim()) {
      toast.error('업체명을 입력해주세요');
      return;
    }
    await createCustomer.mutateAsync({
      name: form.name.trim() || form.companyName.trim(),
      companyName: form.companyName.trim(),
      phone: form.phone.trim() || undefined,
      email: form.email.trim() || undefined,
      address: form.address.trim() || undefined,
      memo: form.memo.trim() || undefined,
      customerType: 'supplier',
    });
    onClose();
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl w-full max-w-md mx-4 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-100">
          <h2 className="text-sm font-bold text-indigo-black">매입처 추가</h2>
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center">
            <X size={16} />
          </button>
        </div>
        <div className="px-5 py-4 space-y-3 max-h-[60vh] overflow-y-auto">
          <p className="text-[10px] text-neutral-400 font-semibold uppercase tracking-wider">기본 정보</p>
          <div className="grid grid-cols-2 gap-3">
            <div className="col-span-2">
              <label className="text-xs text-neutral-500">업체명 *</label>
              <input type="text" value={form.companyName} onChange={(e) => setForm({ ...form, companyName: e.target.value })}
                placeholder="매입처 업체명" autoFocus
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">사업자등록번호</label>
              <input type="text" value={form.businessNumber} onChange={(e) => setForm({ ...form, businessNumber: e.target.value })}
                placeholder="000-00-00000"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">대표자명</label>
              <input type="text" value={form.representative} onChange={(e) => setForm({ ...form, representative: e.target.value })}
                placeholder="대표자"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">업태</label>
              <input type="text" value={form.businessType} onChange={(e) => setForm({ ...form, businessType: e.target.value })}
                placeholder="제조, 도소매 등"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">종목</label>
              <input type="text" value={form.businessCategory} onChange={(e) => setForm({ ...form, businessCategory: e.target.value })}
                placeholder="미용기기, 포장재 등"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
          </div>

          <p className="text-[10px] text-neutral-400 font-semibold uppercase tracking-wider pt-2">연락처</p>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-neutral-500">담당자명</label>
              <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="담당자"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">전화번호</label>
              <input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
                placeholder="010-0000-0000"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">이메일</label>
              <input type="email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })}
                placeholder="email@example.com"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">연락 경로</label>
              <input type="text" value={form.contactChannel} onChange={(e) => setForm({ ...form, contactChannel: e.target.value })}
                placeholder="카톡, 전화, 이메일 등"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
          </div>

          <div>
            <label className="text-xs text-neutral-500">사업장 주소</label>
            <input type="text" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })}
              placeholder="사업장 주소"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
          </div>

          <div>
            <label className="text-xs text-neutral-500">메모</label>
            <input type="text" value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })}
              placeholder="거래 조건, 결제 방식 등"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
          </div>
        </div>
        <div className="px-5 py-4 border-t border-neutral-100 flex gap-2">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button className="flex-1" disabled={!form.companyName.trim() || createCustomer.isPending} onClick={handleSubmit}>
            {createCustomer.isPending ? '등록 중...' : '매입처 등록'}
          </Button>
        </div>
      </div>
    </div>
  );
}
