'use client';

import { useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { Button } from '@/components/ui/button';
import { X } from 'lucide-react';
import toast from 'react-hot-toast';
import { useSetting } from '@/hooks/use-settings';
import { TagSelector } from '@/components/shared/tag-selector';

const TYPE_OPTIONS = [
  { value: 'retail', label: '일반' },
  { value: 'dealer', label: '딜러' },
  { value: 'academy', label: '아카데미' },
];

interface CustomerCreateData {
  name: string;
  phone: string;
  customer_type: string;
  postcode: string;
  address_road: string;
  address_detail: string;
  company_name: string;
  memo: string;
  tags: string[];
}

interface Props {
  open: boolean;
  onClose: () => void;
  onCreated: (customer: { id: string; name: string; phone: string; customer_type: string }) => void;
}

export function CustomerCreateModal({ open, onClose, onCreated }: Props) {
  const [form, setForm] = useState<CustomerCreateData>({
    name: '', phone: '', customer_type: 'retail',
    postcode: '', address_road: '', address_detail: '',
    company_name: '', memo: '', tags: [],
  });
  const [saving, setSaving] = useState(false);
  const queryClient = useQueryClient();
  const availableTags = useSetting<string[]>('customer.tags', []);

  if (!open) return null;

  function openPostcode() {
    const w = window as unknown as Record<string, unknown>;
    if (!w.daum) {
      const s = document.createElement('script');
      s.src = 'https://t1.daumcdn.net/mapjsapi/bundle/postcode/prod/postcode.v2.js';
      s.onload = () => doOpen();
      document.head.appendChild(s);
    } else { doOpen(); }
    function doOpen() {
      new ((window as unknown as Record<string, unknown> & { daum: { Postcode: new (o: Record<string, unknown>) => { open: () => void } } }).daum.Postcode)({
        oncomplete: (data: { zonecode: string; roadAddress: string; jibunAddress: string }) => {
          setForm((prev) => ({ ...prev, postcode: data.zonecode, address_road: data.roadAddress || data.jibunAddress }));
        },
      }).open();
    }
  }

  async function handleSave() {
    if (!form.name.trim()) return toast.error('이름을 입력해주세요');
    if (!form.phone.trim()) return toast.error('연락처를 입력해주세요');

    setSaving(true);
    try {
      const res = await fetch('/api/customers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: form.name.trim(),
          phone: form.phone.trim(),
          customer_type: form.customer_type,
          postcode: form.postcode || undefined,
          address_road: form.address_road || undefined,
          address_detail: form.address_detail || undefined,
          company_name: form.company_name || undefined,
          memo: form.memo || undefined,
          tags: form.tags.length > 0 ? form.tags : undefined,
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error));
      }
      const data = await res.json();
      toast.success('고객 등록 완료');
      queryClient.invalidateQueries({ queryKey: ['customers'] });
      onCreated(data.customer || data);
      onClose();
      setForm({ name: '', phone: '', customer_type: 'retail', postcode: '', address_road: '', address_detail: '', company_name: '', memo: '', tags: [] });
    } catch (err) {
      toast.error(err instanceof Error ? err.message : '등록 실패');
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-[95vw] max-w-[480px] max-h-[90vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">고객 등록</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-3">
          {/* 이름 + 전화 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">이름 <span className="text-red-500">*</span></label>
              <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="홍길동" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">연락처 <span className="text-red-500">*</span></label>
              <input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
                placeholder="010-1234-5678" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
          </div>

          {/* 유형 + 매장명 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">유형</label>
              <div className="flex gap-1.5">
                {TYPE_OPTIONS.map((t) => (
                  <button key={t.value} type="button" onClick={() => setForm({ ...form, customer_type: t.value })}
                    className={`flex-1 py-1.5 text-xs rounded-md border transition ${form.customer_type === t.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}>
                    {t.label}
                  </button>
                ))}
              </div>
            </div>
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">매장/회사명</label>
              <input type="text" value={form.company_name} onChange={(e) => setForm({ ...form, company_name: e.target.value })}
                placeholder="선택 입력" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
          </div>

          {/* 주소 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">주소</label>
            <div className="flex gap-2 mb-2">
              <input type="text" value={form.postcode} readOnly placeholder="우편번호"
                className="w-24 h-9 px-3 rounded-lg border border-neutral-200 bg-neutral-50 text-sm text-neutral-600" />
              <button type="button" onClick={openPostcode}
                className="h-9 px-3 rounded-lg bg-neutral-900 text-white text-xs font-medium">주소검색</button>
            </div>
            <input type="text" value={form.address_road} readOnly placeholder="도로명 주소 (주소검색으로 입력)"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-neutral-50 text-sm text-neutral-600" />
            <input type="text" value={form.address_detail} onChange={(e) => setForm({ ...form, address_detail: e.target.value })}
              placeholder="상세 주소 (동/호수)"
              className="w-full h-9 px-3 mt-2 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
          </div>

          {/* 메모 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">메모</label>
            <input type="text" value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })}
              placeholder="선택 입력" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
          </div>

          {/* 태그 */}
          {availableTags.length > 0 && (
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">태그</label>
              <TagSelector availableTags={availableTags} selectedTags={form.tags} onChange={(tags) => setForm({ ...form, tags })} />
            </div>
          )}
        </div>

        <div className="px-5 py-3 border-t border-neutral-200 flex gap-2">
          <button onClick={onClose} className="flex-1 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600">취소</button>
          <Button className="flex-1" onClick={handleSave} disabled={saving}>
            {saving ? '등록 중...' : '고객 등록'}
          </Button>
        </div>
      </div>
    </div>
  );
}
