'use client';

import { useState } from 'react';
import { useQuery } from '@tanstack/react-query';
import { MapPin, Pencil, Search } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Modal } from '@/components/ui/modal';
import { DaumPostcodeButton } from '@/components/shared/daum-postcode-button';
import { useUpdateConsultationAddress, useDecideCustomerAddress } from '@/hooks/use-consultations';
import { createClient } from '@/lib/supabase/client';

/**
 * 출장 상담 상세의 주소 줄 + [수정] + 고객 주소 확인 창 (2026-11-22, 작업실 t-261121-field-address · 사장님 지시)
 *
 * 1) 주소 고치기 — 도로명은 주소 검색으로만, 상세 주소는 직접 입력. 저장 때 고객 정보 주소와 다르면
 *    「고객 주소도 바꿀까요?」 창을 띄운다([이 건만] / [고객 주소도 바꾸기]). 고객 주소가 비어 있으면 묻지 않고 채운다.
 * 2) 접수 주소 ≠ 고객 정보 주소 — 출장 접수는 기존 고객 주소를 덮어쓰지 않는다(lib/customer/match-or-create keepExistingAddress).
 *    상세를 열면 확인 창이 한 번 뜨고, 고른 것은 상담 건(gas_raw.address_decision)에 남아 다시 묻지 않는다.
 *    새 고객은 접수 때 주소가 그대로 들어가 이 창이 뜨지 않는다.
 * 서버가 지도 좌표 · 구글 달력 장소를 새 주소로 맞추고 이력에 「내용 수정」으로 남긴다. 알림톡은 나가지 않는다.
 */
const FIELD = 'w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-400';

const norm = (road?: string | null, detail?: string | null) => `${(road || '').trim()} ${(detail || '').trim()}`.replace(/\s+/g, ' ').trim();

function useCustomerAddress(customerId: string | null) {
  return useQuery({
    queryKey: ['customer-address', customerId],
    enabled: !!customerId,
    queryFn: async () => {
      const { data } = await createClient().from('customers').select('postcode, address_road, address_detail').eq('id', customerId as string).maybeSingle();
      return (data || null) as { postcode: string | null; address_road: string | null; address_detail: string | null } | null;
    },
  });
}

export function ConsultationAddressLine({ consultationId, customerId, postcode, road, detail, decided }: {
  consultationId: string;
  customerId: string | null;
  postcode: string | null;
  road: string | null;
  detail: string | null;
  /** 접수 주소 ↔ 고객 주소에 대해 이미 고른 적이 있는지(gas_raw.address_decision) */
  decided: boolean;
}) {
  const save = useUpdateConsultationAddress();
  const decide = useDecideCustomerAddress();
  const { data: cust, isFetched } = useCustomerAddress(customerId);
  const [editing, setEditing] = useState(false);
  const [form, setForm] = useState({ postcode: postcode || '', road: road || '', detail: detail || '' });
  const [askOnSave, setAskOnSave] = useState(false);     // 고친 주소를 고객 주소로도?
  const [mismatchClosed, setMismatchClosed] = useState(false); // 이번에 열린 화면에서 닫았는지

  const custAddr = norm(cust?.address_road, cust?.address_detail);
  const custHasAddr = !!(cust?.address_road || '').trim();

  // 2) 상세를 열었을 때 접수 주소가 고객 정보 주소와 다르고 아직 고른 적이 없으면 묻는다(닫으면 이 화면에서는 다시 안 뜸)
  const askMismatch = isFetched && !decided && !!customerId && custHasAddr && !!road && norm(road, detail) !== custAddr && !mismatchClosed;
  const setAskMismatch = (v: boolean) => { if (!v) setMismatchClosed(true); };

  const open = () => { setForm({ postcode: postcode || '', road: road || '', detail: detail || '' }); setEditing(true); };
  const changed = form.road.trim() !== (road || '').trim() || form.detail.trim() !== (detail || '').trim();
  const doSave = (alsoCustomer: boolean) => save.mutate(
    { id: consultationId, postcode: form.postcode, road: form.road.trim(), detail: form.detail.trim(), alsoCustomer },
    { onSuccess: () => { setEditing(false); setAskOnSave(false); } },
  );
  const onSave = () => {
    if (!customerId) return doSave(false);
    if (!custHasAddr) return doSave(true);                                   // 고객 주소가 비어 있으면 묻지 않고 채운다
    if (norm(form.road, form.detail) === custAddr) return doSave(false);     // 이미 같은 주소
    setAskOnSave(true);
  };

  return (
    <>
      {!editing ? (
        <div className="flex items-start gap-2 text-sm" data-address-line>
          <MapPin size={14} className="text-neutral-400 mt-0.5 shrink-0" />
          <span className="min-w-0 flex-1 break-keep">{road ? `${road} ${detail || ''}` : <span className="text-neutral-400">주소 없음</span>}</span>
          <button type="button" onClick={open} data-address-edit className="shrink-0 inline-flex items-center gap-1 text-xs text-neutral-500 hover:text-stone-900">
            <Pencil size={11} />{road ? '수정' : '주소 넣기'}
          </button>
        </div>
      ) : (
        <div className="rounded-lg border border-neutral-200 bg-stone-50 p-3 space-y-2" data-address-form>
          <p className="text-xs font-semibold text-stone-800">주소 고치기</p>
          <div className="flex gap-2">
            <input type="text" value={form.road} readOnly placeholder="주소 검색을 눌러 주세요" className={`${FIELD} bg-neutral-50`} aria-label="도로명 주소" />
            <DaumPostcodeButton
              onSelected={({ zonecode, roadAddress }) => setForm((f) => ({ ...f, postcode: zonecode, road: roadAddress, detail: '' }))}
              className="shrink-0 h-9 px-3 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 inline-flex items-center gap-1"
            >
              <Search size={12} />주소 검색
            </DaumPostcodeButton>
          </div>
          <input type="text" value={form.detail} maxLength={200} onChange={(e) => setForm({ ...form, detail: e.target.value })} placeholder="상세 주소 (동 · 호수 · 상호)" className={FIELD} aria-label="상세 주소" />
          <p className="text-[11px] text-neutral-400">저장하면 지도 · 구글 달력이 새 주소로 바뀝니다. 고객에게 알림톡은 나가지 않습니다.</p>
          <div className="flex justify-end gap-2">
            <Button variant="ghost" size="sm" onClick={() => setEditing(false)} disabled={save.isPending} data-safe>닫기</Button>
            <Button size="sm" disabled={!form.road.trim() || !changed || save.isPending} onClick={onSave}>
              {save.isPending ? '저장 중...' : '주소 저장'}
            </Button>
          </div>
        </div>
      )}

      {/* 1) 고친 주소 — 고객 정보 주소도 바꿀지 */}
      <Modal open={askOnSave} onClose={() => setAskOnSave(false)} title="고객 주소도 바꿀까요?" className="max-w-sm">
        <div className="space-y-3 text-sm" data-ask-customer-address>
          <AddrCompare savedLabel="고객 정보 주소" saved={custAddr} nextLabel="고친 주소" next={norm(form.road, form.detail)} />
          <div className="flex justify-end gap-2">
            <Button variant="ghost" size="sm" onClick={() => doSave(false)} disabled={save.isPending} data-safe>이 건만 고치기</Button>
            <Button size="sm" onClick={() => doSave(true)} disabled={save.isPending}>고객 주소도 바꾸기</Button>
          </div>
        </div>
      </Modal>

      {/* 2) 접수 주소가 고객 정보 주소와 다름 */}
      <Modal open={askMismatch} onClose={() => setAskMismatch(false)} title="접수 주소가 저장된 주소와 다릅니다" className="max-w-sm">
        <div className="space-y-3 text-sm" data-ask-mismatch>
          <p className="text-xs text-neutral-500">고객이 출장 요청에 적은 주소가 고객 정보에 저장된 주소와 다릅니다. 고객 주소를 바꿀까요?</p>
          <AddrCompare savedLabel="저장된 주소" saved={custAddr} nextLabel="이번 접수 주소" next={norm(road, detail)} />
          <div className="flex justify-end gap-2">
            <Button variant="ghost" size="sm" disabled={decide.isPending} data-safe
              onClick={() => decide.mutate({ id: consultationId, decision: 'keep' }, { onSuccess: () => setAskMismatch(false) })}>그대로 두기</Button>
            <Button size="sm" disabled={decide.isPending}
              onClick={() => decide.mutate({ id: consultationId, decision: 'apply' }, { onSuccess: () => setAskMismatch(false) })}>고객 주소로 바꾸기</Button>
          </div>
          <p className="text-[11px] text-neutral-400">어느 쪽을 골라도 이 출장 건의 주소는 그대로입니다. 닫으면 다음에 열 때 다시 묻습니다.</p>
        </div>
      </Modal>
    </>
  );
}

function AddrCompare({ savedLabel, saved, nextLabel, next }: { savedLabel: string; saved: string; nextLabel: string; next: string }) {
  return (
    <div className="rounded-lg border border-neutral-200 divide-y divide-neutral-100 text-xs">
      <div className="px-3 py-2"><p className="text-neutral-400">{savedLabel}</p><p className="text-neutral-700 break-keep">{saved || '(없음)'}</p></div>
      <div className="px-3 py-2"><p className="text-neutral-400">{nextLabel}</p><p className="font-semibold text-stone-900 break-keep">{next || '(없음)'}</p></div>
    </div>
  );
}
