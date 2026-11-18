'use client';

import { useState } from 'react';
import toast from 'react-hot-toast';
import { Info, AlertTriangle, Copy, Truck, RotateCcw, Pencil, Search } from 'lucide-react';
import { CustomerAutocomplete, type SelectedCustomer } from '@/components/shared/customer-autocomplete';
import { useCreateManualInvoice } from '@/hooks/use-manual-invoices';
import type { ManualInvoice } from '@/lib/supabase/types';
import { formatPhone } from '@/lib/utils/format';
import { DaumPostcodeButton } from '@/components/shared/daum-postcode-button';

/** 받는 곳 (이번 송장에 실제로 찍히는 값) */
interface Receiver { name: string; phone: string; postcode: string; road: string; detail: string }
const fromCustomer = (c: SelectedCustomer): Receiver => ({ name: c.name || '', phone: c.phone || '', postcode: c.postcode || '', road: c.address_road || '', detail: c.address_detail || '' });
const FIELD = 'h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40';
const INPUT = `w-full ${FIELD}`;

export function QuickInvoiceForm() {
  const [selectedCustomer, setSelectedCustomer] = useState<SelectedCustomer | null>(null);
  const [goodsName, setGoodsName] = useState('');
  const [deliveryMessage, setDeliveryMessage] = useState('');
  const [lastIssued, setLastIssued] = useState<ManualInvoice | null>(null);

  const createInvoice = useCreateManualInvoice();

  // 2026-11-18 받는 곳을 이번 송장에서만 다르게 적을 수 있다 — 전엔 고객에 저장된 주소로만 나가서, 다른 곳으로 보낼 때 고객 정보를 먼저 고쳐야 했다
  const [recv, setRecv] = useState<Receiver>({ name: '', phone: '', postcode: '', road: '', detail: '' });
  const [editing, setEditing] = useState(false);
  const [saveAddr, setSaveAddr] = useState(false);

  const saved = selectedCustomer ? fromCustomer(selectedCustomer) : null;
  const changed = !!saved && (Object.keys(saved) as (keyof Receiver)[]).some((k) => saved[k].trim() !== recv[k].trim());
  const addrChanged = !!saved && (saved.postcode.trim() !== recv.postcode.trim() || saved.road.trim() !== recv.road.trim() || saved.detail.trim() !== recv.detail.trim());
  const missing = [
    !recv.name.trim() && '받는 사람',
    recv.phone.replace(/\D/g, '').length < 9 && '연락처',
    (!recv.postcode.trim() || !recv.road.trim()) && '주소',
  ].filter(Boolean) as string[];

  function pickCustomer(c: SelectedCustomer) {
    const r = fromCustomer(c);
    setSelectedCustomer(c);
    setRecv(r);
    setSaveAddr(false);
    // 저장된 주소 · 연락처가 비어 있으면 바로 적을 수 있게 펼친다 (전엔 고객 화면으로 가서 채워야 했다)
    setEditing(!r.phone || !r.postcode || !r.road);
  }

  const canIssue =
    !!selectedCustomer &&
    missing.length === 0 &&
    goodsName.trim().length > 0 &&
    goodsName.trim().length <= 50;

  function reset() {
    setSelectedCustomer(null);
    setEditing(false);
    setSaveAddr(false);
    setGoodsName('');
    setDeliveryMessage('');
    setLastIssued(null);
  }

  async function handleIssue() {
    if (!selectedCustomer || !canIssue) return;
    try {
      const res = await createInvoice.mutateAsync({
        customer_id: selectedCustomer.id,
        goods_name: goodsName.trim(),
        delivery_message: deliveryMessage.trim() || undefined,
        // 바꾼 것이 없으면 보내지 않는다 — 서버가 고객에 저장된 값을 그대로 쓴다
        ...(changed ? {
          receiver: { name: recv.name.trim(), phone: recv.phone.trim(), postcode: recv.postcode.trim(), address_road: recv.road.trim(), address_detail: recv.detail.trim() || undefined },
          save_address_to_customer: saveAddr && addrChanged,
        } : {}),
      });
      if (res.invoice) setLastIssued(res.invoice);
    } catch {
      // 토스트는 hook의 onError가 처리
    }
  }

  async function copyInvoiceNumber(num: string) {
    try {
      await navigator.clipboard.writeText(num);
      toast.success('송장번호 복사 완료');
    } catch {
      toast.error('복사 실패');
    }
  }

  // ── 발급 결과 화면 ──
  if (lastIssued) {
    return (
      <div className="space-y-4">
        <div className="rounded-xl border border-green-200 bg-green-50 p-5">
          <div className="flex items-center gap-2 mb-3">
            <Truck size={18} className="text-green-700" />
            <span className="text-sm font-bold text-green-800">송장 발급 완료</span>
          </div>
          <div className="flex items-center justify-between gap-3 bg-white rounded-lg border border-green-200 px-4 py-3">
            <span className="text-3xl font-bold font-mono tracking-wider text-indigo-black select-all">
              {lastIssued.invoice_number}
            </span>
            <button
              type="button"
              onClick={() => copyInvoiceNumber(lastIssued.invoice_number)}
              className="flex items-center gap-1.5 px-3 py-2 rounded-lg bg-indigo-black text-cream text-xs font-semibold hover:bg-indigo-black/85 transition shrink-0"
            >
              <Copy size={13} />
              복사
            </button>
          </div>
          <p className="text-xs text-green-700 mt-3">
            ALPS 별도 PC에서 이 송장번호로 조회 → 출력하세요.
          </p>
          <div className="mt-3 text-[11px] text-neutral-500 space-y-0.5">
            <p>받는 사람: {lastIssued.customer_name} {lastIssued.customer_phone && `(${formatPhone(lastIssued.customer_phone)})`}</p>
            <p>주소: {[lastIssued.receiver_postcode, lastIssued.receiver_address_road, lastIssued.receiver_address_detail].filter(Boolean).join(' ')}</p>
            <p>품목: {lastIssued.goods_name}</p>
          </div>
        </div>

        <button
          type="button"
          onClick={reset}
          className="w-full h-10 rounded-lg border border-neutral-200 text-sm font-semibold text-neutral-700 hover:bg-neutral-50 flex items-center justify-center gap-2"
        >
          <RotateCcw size={14} />
          한 건 더 발급
        </button>
      </div>
    );
  }

  // ── 발급 폼 ──
  return (
    <div className="space-y-4">
      {/* 안내 박스 */}
      <div className="flex items-start gap-2 rounded-lg bg-warm-ivory border border-neutral-200 px-3 py-2.5">
        <Info size={14} className="text-neutral-500 mt-0.5 shrink-0" />
        <p className="text-xs text-neutral-600 leading-relaxed">
          판매와 무관한 송장입니다. 매출 통계에 반영되지 않습니다. 샘플 발송·1회성 출고·간단 AS·거래처 출고용으로 사용하세요.
        </p>
      </div>

      {/* 1. 고객 선택 */}
      <div>
        <label className="block text-xs font-semibold text-neutral-600 mb-1.5">고객 *</label>
        <CustomerAutocomplete
          selectedCustomer={selectedCustomer}
          onSelect={pickCustomer}
          onClear={() => { setSelectedCustomer(null); setEditing(false); setSaveAddr(false); }}
        />

        {/* 받는 곳 — 보기 */}
        {selectedCustomer && !editing && (
          <div className="mt-2 rounded-lg bg-terracotta/5 border border-terracotta/20 px-3 py-2.5 text-xs flex items-start gap-3">
            <div className="flex-1 min-w-0 space-y-0.5">
              <p className="text-neutral-600">
                <span className="text-neutral-400">받는 사람</span>{' '}
                <span className="font-semibold text-indigo-black">{recv.name}</span>
                {recv.phone && <span className="ml-2 text-neutral-500">{formatPhone(recv.phone)}</span>}
              </p>
              <p className="text-neutral-600 break-keep">
                <span className="text-neutral-400">주소</span>{' '}
                {[recv.postcode, recv.road, recv.detail].filter(Boolean).join(' ')}
              </p>
              {changed && <p className="text-[11px] font-semibold text-amber-700">이번 송장만 이 정보로 나갑니다{saveAddr && addrChanged ? ' · 주소는 고객 정보에도 저장' : ''}</p>}
            </div>
            <button type="button" onClick={() => setEditing(true)}
              className="shrink-0 h-8 px-2.5 rounded-lg border border-neutral-300 bg-white text-[11px] font-semibold text-neutral-700 hover:bg-neutral-50 inline-flex items-center gap-1">
              <Pencil size={12} />받는 곳 바꾸기
            </button>
          </div>
        )}

        {/* 받는 곳 — 고치기 */}
        {selectedCustomer && editing && (
          <div className="mt-2 rounded-lg border border-neutral-200 bg-white px-3 py-3 space-y-2">
            <div className="flex items-center justify-between gap-2">
              <p className="text-xs font-bold text-neutral-700">받는 곳 <span className="font-normal text-neutral-400">(이번 송장에 찍히는 정보)</span></p>
              {changed && (
                <button type="button" onClick={() => { if (saved) setRecv(saved); setSaveAddr(false); }}
                  className="text-[11px] text-neutral-500 underline hover:text-neutral-700">저장된 정보로 되돌리기</button>
              )}
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <div>
                <label className="block text-[11px] text-neutral-500 mb-0.5">받는 사람</label>
                <input type="text" value={recv.name} maxLength={40} onChange={(e) => setRecv({ ...recv, name: e.target.value })} placeholder="이름" className={INPUT} aria-label="받는 사람" />
              </div>
              <div>
                <label className="block text-[11px] text-neutral-500 mb-0.5">연락처</label>
                <input type="tel" value={recv.phone} maxLength={20} onChange={(e) => setRecv({ ...recv, phone: e.target.value })} placeholder="010-0000-0000" className={INPUT} aria-label="받는 사람 연락처" />
              </div>
            </div>
            <div>
              <label className="block text-[11px] text-neutral-500 mb-0.5">주소</label>
              {/* 우편번호 + 검색 버튼 한 줄, 도로명은 아래 줄 전체 — 휴대폰 폭에서도 넘치지 않게 */}
              <div className="flex gap-2">
                <input type="text" value={recv.postcode} readOnly placeholder="우편번호" className={`${FIELD} w-24 shrink-0 bg-neutral-50`} aria-label="우편번호" />
                <DaumPostcodeButton
                  onSelected={({ zonecode, roadAddress }) => setRecv((r) => ({ ...r, postcode: zonecode, road: roadAddress, detail: '' }))}
                  className="shrink-0 h-9 px-3 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 inline-flex items-center gap-1"
                >
                  <Search size={12} />주소 검색
                </DaumPostcodeButton>
              </div>
              <input type="text" value={recv.road} readOnly placeholder="주소 검색을 눌러 주세요" className={`${INPUT} mt-2 bg-neutral-50`} aria-label="도로명 주소" />
              <input type="text" value={recv.detail} maxLength={200} onChange={(e) => setRecv({ ...recv, detail: e.target.value })} placeholder="상세 주소 (동 · 호수 · 상호)" className={`${INPUT} mt-2`} aria-label="상세 주소" />
            </div>
            {addrChanged && (
              <label className="flex items-start gap-2 text-xs text-neutral-700 cursor-pointer py-1">
                <input type="checkbox" checked={saveAddr} onChange={(e) => setSaveAddr(e.target.checked)} className="w-4 h-4 mt-0.5 shrink-0" />
                <span>이 주소를 고객 정보에도 저장<span className="block text-[11px] text-neutral-400">끄면 이번 송장에만 씁니다</span></span>
              </label>
            )}
            {missing.length > 0 ? (
              <p className="text-[11px] text-red-600 flex items-center gap-1"><AlertTriangle size={12} />빠진 항목: {missing.join(' · ')}</p>
            ) : (
              <button type="button" onClick={() => setEditing(false)}
                className="h-8 px-3 rounded-lg border border-neutral-300 bg-white text-xs font-semibold text-neutral-700 hover:bg-neutral-50">확인</button>
            )}
          </div>
        )}
      </div>

      {/* 2. 품목명 */}
      <div>
        <div className="flex items-center justify-between mb-1.5">
          <label className="block text-xs font-semibold text-neutral-600">
            품목명 * <span className="text-neutral-400 font-normal">(송장에 인쇄됩니다)</span>
          </label>
          <span className={`text-[11px] ${goodsName.length > 50 ? 'text-red-500' : 'text-neutral-400'}`}>
            {goodsName.length}/50
          </span>
        </div>
        <input
          type="text"
          value={goodsName}
          onChange={(e) => setGoodsName(e.target.value)}
          maxLength={50}
          placeholder="예: 샘플 가위 1점, 수리 후 반환 가위, 거래처 출고품"
          className="w-full h-10 px-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
        />
      </div>

      {/* 3. 배송 메시지 (선택) */}
      <div>
        <label className="block text-xs font-semibold text-neutral-600 mb-1.5">
          배송 메시지 <span className="text-neutral-400 font-normal">(선택)</span>
        </label>
        <input
          type="text"
          value={deliveryMessage}
          onChange={(e) => setDeliveryMessage(e.target.value)}
          maxLength={100}
          placeholder="예: 부재시 경비실에 맡겨주세요"
          className="w-full h-10 px-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
        />
      </div>

      {/* 4. 발급 버튼 */}
      <button
        type="button"
        onClick={handleIssue}
        disabled={!canIssue || createInvoice.isPending}
        className="w-full h-12 rounded-lg bg-terracotta text-cream text-sm font-bold hover:bg-terracotta/90 disabled:opacity-40 disabled:cursor-not-allowed transition flex items-center justify-center gap-2"
      >
        <Truck size={16} />
        {createInvoice.isPending ? '발급 중...' : 'ALPS 송장 발급'}
      </button>
    </div>
  );
}
