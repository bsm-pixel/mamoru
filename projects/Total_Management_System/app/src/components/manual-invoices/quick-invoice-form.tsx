'use client';

import { useState } from 'react';
import Link from 'next/link';
import toast from 'react-hot-toast';
import { Info, AlertTriangle, Copy, Truck, RotateCcw } from 'lucide-react';
import { CustomerAutocomplete, type SelectedCustomer } from '@/components/shared/customer-autocomplete';
import { useCreateManualInvoice } from '@/hooks/use-manual-invoices';
import type { ManualInvoice } from '@/lib/supabase/types';
import { formatPhone } from '@/lib/utils/format';

export function QuickInvoiceForm() {
  const [selectedCustomer, setSelectedCustomer] = useState<SelectedCustomer | null>(null);
  const [goodsName, setGoodsName] = useState('');
  const [deliveryMessage, setDeliveryMessage] = useState('');
  const [lastIssued, setLastIssued] = useState<ManualInvoice | null>(null);

  const createInvoice = useCreateManualInvoice();

  const addressMissing = !!selectedCustomer && (!selectedCustomer.postcode || !selectedCustomer.address_road);
  const phoneMissing = !!selectedCustomer && !selectedCustomer.phone;
  const customerIncomplete = addressMissing || phoneMissing;

  const canIssue =
    !!selectedCustomer &&
    !customerIncomplete &&
    goodsName.trim().length > 0 &&
    goodsName.trim().length <= 50;

  function reset() {
    setSelectedCustomer(null);
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
          onSelect={(c) => setSelectedCustomer(c)}
          onClear={() => setSelectedCustomer(null)}
        />

        {/* 고객 정보 미리보기 */}
        {selectedCustomer && !customerIncomplete && (
          <div className="mt-2 rounded-lg bg-terracotta/5 border border-terracotta/20 px-3 py-2.5 text-xs space-y-0.5">
            <p className="text-neutral-600">
              <span className="text-neutral-400">받는 사람</span>{' '}
              <span className="font-semibold text-indigo-black">{selectedCustomer.name}</span>
              {selectedCustomer.phone && <span className="ml-2 text-neutral-500">{formatPhone(selectedCustomer.phone)}</span>}
            </p>
            <p className="text-neutral-600">
              <span className="text-neutral-400">주소</span>{' '}
              {[selectedCustomer.postcode, selectedCustomer.address_road, selectedCustomer.address_detail].filter(Boolean).join(' ')}
            </p>
          </div>
        )}

        {/* 차단 메시지 */}
        {selectedCustomer && customerIncomplete && (
          <div className="mt-2 rounded-lg bg-red-50 border border-red-200 px-3 py-2.5">
            <div className="flex items-start gap-2">
              <AlertTriangle size={14} className="text-red-500 mt-0.5 shrink-0" />
              <div className="flex-1 text-xs space-y-1">
                <p className="font-semibold text-red-700">
                  {addressMissing && phoneMissing
                    ? '고객 주소·연락처가 모두 비어 있습니다.'
                    : addressMissing
                    ? '고객 주소(우편번호+도로명)가 비어 있습니다.'
                    : '고객 연락처가 비어 있습니다.'}
                </p>
                <p className="text-red-600">
                  송장 발급을 위해 먼저 고객 정보를 보강해주세요.
                </p>
                <Link
                  href={`/customers/${selectedCustomer.id}`}
                  className="inline-block text-red-700 underline font-semibold hover:text-red-800"
                >
                  고객 정보 페이지로 이동 →
                </Link>
              </div>
            </div>
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
