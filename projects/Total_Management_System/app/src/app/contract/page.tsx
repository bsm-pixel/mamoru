'use client';

import { useState, useMemo } from 'react';
import { Button } from '@/components/ui/button';
import { SignatureCanvas } from '@/components/contracts/signature-canvas';
import { ProductPickerModal } from '@/components/contracts/product-picker-modal';
import { useCreateContract } from '@/hooks/use-contracts';
import { formatKRW } from '@/lib/utils/format';
import { Plus, X, CheckCircle2, RotateCcw } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';

/* ── 제품 행 ────────────────────────── */
interface ProductRow {
  id: string;
  product: Product | null;
  quantity: number;
}

/* ── 법적 문구 (종이 계약서 원문) ────── */
const LEGAL_NOTICE = `마모루는 도움을 드리는 회사입니다.

저희는 판매목적의 영리목적이 아니라고 생각하며
여러분의 행복은 고객님의 부분이 될 수 있다 생각합니다.

그러기에 직접 확인의 이유로 분해나 발열이 발생 시
영역 받지않은 개인이나 지점으로 수리를 맡기는 행위는 안됩니다.

기본적인 관리, 재연마 방법은
마모루에서 교육(무상) 해드리겠습니다.`;

const CAUTION_NOTICE = `제품수령일 기준 5일 이내 제품 교환 및 반품 가능합니다.
※ 어떠한 명분이든 교환 등 환불은 불가합니다 ※
기간 내라도, 문구 각인이 진행된 경우
고객님의 취급 부주의로 인한 제품의 사용성이 손실된 경우
5일이 지난 이후의 반납은 교환 및 환불이 어렵습니다.`;

export default function ContractStandalonePage() {
  const createContract = useCreateContract();


  /* ── 완료 상태 ── */
  const [saved, setSaved] = useState(false);

  /* ── 매장 정보 ── */
  const [shopName, setShopName] = useState('');
  const [shopAddress, setShopAddress] = useState('');

  /* ── 제품 ── */
  const [rows, setRows] = useState<ProductRow[]>([{ id: crypto.randomUUID(), product: null, quantity: 1 }]);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [activeRowId, setActiveRowId] = useState<string | null>(null);

  /* ── 결제 ── */
  const [paymentMethod, setPaymentMethod] = useState('transfer');
  const [installment, setInstallment] = useState(0);
  const [discount, setDiscount] = useState(0);
  const [depositAmount, setDepositAmount] = useState(0);
  const [memo, setMemo] = useState('');

  /* ── 서명 ── */
  const [buyerSignature, setBuyerSignature] = useState('');
  const [sellerSignature, setSellerSignature] = useState('');

  /* ── 계산 ── */
  const totalAmount = useMemo(() =>
    rows.reduce((s, r) => s + (r.product?.price || 0) * r.quantity, 0), [rows]);
  const finalAmount = totalAmount - discount;
  const balanceAmount = Math.max(0, finalAmount - depositAmount);

  /* ── 날짜 ── */
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');

  /* ── 제품 모달 ── */
  function openPicker(rowId: string) {
    setActiveRowId(rowId);
    setPickerOpen(true);
  }

  function handleProductSelect(product: Product) {
    if (!activeRowId) return;
    setRows((prev) =>
      prev.map((r) => r.id === activeRowId ? { ...r, product, id: product.id + '_' + Date.now() } : r)
    );
  }

  function addRow() {
    setRows((prev) => [...prev, { id: crypto.randomUUID(), product: null, quantity: 1 }]);
  }

  function removeRow(rowId: string) {
    setRows((prev) => prev.filter((r) => r.id !== rowId));
  }

  function updateQty(rowId: string, qty: number) {
    setRows((prev) => prev.map((r) => r.id === rowId ? { ...r, quantity: Math.max(1, qty) } : r));
  }

  /* ── 제출 ── */
  async function handleSubmit() {
    const filledRows = rows.filter((r) => r.product);
    if (filledRows.length === 0) return;

    await createContract.mutateAsync({
      contract: {
        customer_name: shopName.trim() || '미입력',
        customer_address: shopAddress.trim() || undefined,
        total_amount: totalAmount,
        discount_amount: discount,
        final_amount: finalAmount,
        payment_method: paymentMethod,
        installment_months: installment,
        signature_data: buyerSignature || undefined,
        memo: memo.trim() || undefined,
        delivery_method: 'shipping',
        deposit_amount: depositAmount,
        balance_amount: balanceAmount,
        seller_signature: sellerSignature || undefined,
        shop_name: shopName.trim() || undefined,
        shop_address: shopAddress.trim() || undefined,
      },
      items: filledRows.map((r) => ({
        product_id: r.product!.id,
        product_name: r.product!.name,
        sku: r.product!.sku,
        quantity: r.quantity,
        unit_price: r.product!.price,
        total_price: r.product!.price * r.quantity,
      })),
    });

    setSaved(true);
  }

  /* ── 새 계약서 작성 (폼 초기화) ── */
  function handleReset() {
    setSaved(false);
    setShopName('');
    setShopAddress('');
    setRows([{ id: crypto.randomUUID(), product: null, quantity: 1 }]);
    setPaymentMethod('transfer');
    setInstallment(0);
    setDiscount(0);
    setDepositAmount(0);
    setMemo('');
    setBuyerSignature('');
    setSellerSignature('');
  }

  const canSubmit = rows.some((r) => r.product) && !createContract.isPending;

  /* ── 저장 완료 화면 ── */
  if (saved) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-neutral-50 px-4">
        <div className="text-center space-y-6">
          <CheckCircle2 size={64} className="mx-auto text-green-500" />
          <div>
            <h2 className="text-xl font-bold text-neutral-900">계약서 저장 완료</h2>
            <p className="text-sm text-neutral-500 mt-1">감사합니다</p>
          </div>
          <Button onClick={handleReset} className="gap-2">
            <RotateCcw size={16} />
            새 계약서 작성
          </Button>
        </div>
      </div>
    );
  }

  /* ── 입력 스타일 ── */
  const inputClass = 'border-0 border-b border-neutral-300 bg-transparent px-1 py-1 text-sm focus:outline-none focus:border-neutral-800 w-full';

  return (
    <div className="min-h-screen bg-neutral-100">
      <div className="px-3 py-4">
        <div className="bg-white max-w-[600px] mx-auto border border-neutral-400 rounded-sm shadow-sm">

          {/* ── 헤더 ── */}
          <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-400">
            <span className="text-sm font-bold tracking-wider text-neutral-800">MAMORU</span>
            <h2 className="text-base font-extrabold tracking-[0.3em] text-neutral-900">구 매 계 약 서</h2>
          </div>

          {/* ── 매장 정보 ── */}
          <div className="px-5 py-4 border-b border-neutral-300 space-y-2">
            <div>
              <label className="text-[10px] text-neutral-500">매장명</label>
              <input
                type="text"
                value={shopName}
                onChange={(e) => setShopName(e.target.value)}
                placeholder="매장명"
                className={inputClass}
              />
            </div>
            <div>
              <label className="text-[10px] text-neutral-500">매장주소 (수령지주소)</label>
              <input
                type="text"
                value={shopAddress}
                onChange={(e) => setShopAddress(e.target.value)}
                placeholder="주소"
                className={inputClass}
              />
            </div>
          </div>

          {/* ── 필독 (유의사항) ── */}
          <div className="px-5 py-4 border-b border-neutral-300 bg-neutral-50">
            <p className="text-xs font-bold text-neutral-700 mb-2 text-center">※ 필독 ※</p>
            <p className="text-[11px] text-neutral-600 leading-relaxed whitespace-pre-line">{LEGAL_NOTICE}</p>
            <div className="mt-3 pt-2 border-t border-neutral-200">
              <p className="text-[10px] font-bold text-neutral-700 mb-1">[유의사항]</p>
              <p className="text-[11px] text-neutral-600 leading-relaxed whitespace-pre-line">{CAUTION_NOTICE}</p>
            </div>
          </div>

          {/* ── 결제방식 ── */}
          <div className="px-5 py-4 border-b border-neutral-300 space-y-3">
            <p className="text-xs font-bold text-neutral-700">결제방식</p>
            <div className="flex flex-wrap gap-3">
              {[
                { value: 'transfer', label: '현금/계좌이체' },
                { value: 'card', label: '카드/할부' },
                { value: 'cms', label: 'CMS 자동이체' },
              ].map((opt) => (
                <label key={opt.value} className="flex items-center gap-1.5 text-sm cursor-pointer">
                  <input
                    type="radio"
                    name="payment"
                    checked={paymentMethod === opt.value}
                    onChange={() => setPaymentMethod(opt.value)}
                    className="accent-neutral-800"
                  />
                  {opt.label}
                </label>
              ))}
            </div>

            {paymentMethod === 'card' && (
              <div>
                <label className="text-[10px] text-neutral-500">할부</label>
                <select
                  value={installment}
                  onChange={(e) => setInstallment(parseInt(e.target.value))}
                  className="ml-2 border-b border-neutral-300 bg-transparent text-sm focus:outline-none py-1"
                >
                  <option value={0}>일시불</option>
                  {[2, 3, 4, 5, 6, 9, 12].map((m) => (
                    <option key={m} value={m}>{m}개월</option>
                  ))}
                </select>
              </div>
            )}

            <div className="grid grid-cols-3 gap-3">
              <div>
                <label className="text-[10px] text-neutral-500">총액</label>
                <p className="text-sm font-semibold border-b border-neutral-300 py-1">{formatKRW(finalAmount)}</p>
              </div>
              <div>
                <label className="text-[10px] text-neutral-500">선납금</label>
                <input
                  type="number"
                  value={depositAmount || ''}
                  onChange={(e) => setDepositAmount(parseInt(e.target.value) || 0)}
                  placeholder="0"
                  className={inputClass}
                />
              </div>
              <div>
                <label className="text-[10px] text-neutral-500">잔금</label>
                <p className="text-sm font-semibold border-b border-neutral-300 py-1">{formatKRW(balanceAmount)}</p>
              </div>
            </div>

            <div>
              <label className="text-[10px] text-neutral-500">할인</label>
              <input
                type="number"
                value={discount || ''}
                onChange={(e) => setDiscount(parseInt(e.target.value) || 0)}
                placeholder="0"
                className={inputClass + ' max-w-[140px]'}
              />
            </div>

            <div>
              <label className="text-[10px] text-neutral-500">메모</label>
              <input
                type="text"
                value={memo}
                onChange={(e) => setMemo(e.target.value)}
                placeholder="메모 (선택)"
                className={inputClass}
              />
            </div>
          </div>

          {/* ── 제품 테이블 ── */}
          <div className="px-5 py-4 border-b border-neutral-300">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="border-b border-neutral-400">
                  <th className="text-left py-2 text-xs font-bold text-neutral-700 w-[50%]">품명</th>
                  <th className="text-center py-2 text-xs font-bold text-neutral-700 w-[15%]">수량</th>
                  <th className="text-right py-2 text-xs font-bold text-neutral-700 w-[25%]">금액</th>
                  <th className="w-[10%]"></th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row) => (
                  <tr key={row.id} className="border-b border-neutral-200">
                    <td className="py-3">
                      {row.product ? (
                        <button onClick={() => openPicker(row.id)} className="text-left">
                          <span className="text-sm font-medium text-neutral-800">{row.product.name}</span>
                          <span className="text-[10px] text-neutral-500 block">{row.product.sku}</span>
                        </button>
                      ) : (
                        <button
                          onClick={() => openPicker(row.id)}
                          className="text-sm text-neutral-400 underline underline-offset-4 decoration-dashed py-1"
                        >
                          탭하여 선택
                        </button>
                      )}
                    </td>
                    <td className="text-center py-3">
                      <input
                        type="number"
                        value={row.quantity}
                        onChange={(e) => updateQty(row.id, parseInt(e.target.value) || 1)}
                        className="w-12 text-center border-b border-neutral-300 bg-transparent text-sm focus:outline-none focus:border-neutral-800"
                        min={1}
                      />
                    </td>
                    <td className="text-right py-3 text-sm font-semibold text-neutral-800">
                      {row.product ? formatKRW(row.product.price * row.quantity) : '-'}
                    </td>
                    <td className="text-center py-3">
                      {rows.length > 1 && (
                        <button onClick={() => removeRow(row.id)} className="text-neutral-400 hover:text-red-500">
                          <X size={14} />
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <button
              onClick={addRow}
              className="flex items-center gap-1.5 mt-3 text-xs text-neutral-500 hover:text-terracotta transition"
            >
              <Plus size={14} />
              제품 추가
            </button>

            {totalAmount > 0 && (
              <div className="mt-3 pt-2 border-t border-neutral-300 text-right">
                {discount > 0 && (
                  <p className="text-xs text-neutral-500">할인: -{formatKRW(discount)}</p>
                )}
                <p className="text-sm font-bold text-neutral-900">합계: {formatKRW(finalAmount)}</p>
              </div>
            )}
          </div>

          {/* ── 날짜 + 서명 ── */}
          <div className="px-5 py-4 border-b border-neutral-300 space-y-4">
            <p className="text-center text-sm text-neutral-800 tracking-widest">
              {yyyy} 년 {mm} 월 {dd} 일
            </p>

            <div className="space-y-3">
              <div>
                <p className="text-xs font-bold text-neutral-700 mb-1">구매자 서명</p>
                <SignatureCanvas onSign={setBuyerSignature} width={320} height={120} />
              </div>
              <div>
                <p className="text-xs font-bold text-neutral-700 mb-1">판매자 서명</p>
                <SignatureCanvas onSign={setSellerSignature} width={320} height={120} />
              </div>
            </div>
          </div>

          {/* ── 계좌 안내 ── */}
          <div className="px-5 py-3 bg-neutral-50 text-center">
            <p className="text-[11px] text-neutral-600">
              입금 계좌: <span className="font-semibold">우리은행 1002-439-462514 (백성민)</span>
            </p>
          </div>
        </div>

        {/* ── 저장 버튼 ── */}
        <div className="max-w-[600px] mx-auto mt-4 pb-8">
          <Button
            className="w-full"
            disabled={!canSubmit}
            onClick={handleSubmit}
          >
            {createContract.isPending ? '저장 중...' : `계약서 저장 (${formatKRW(finalAmount)})`}
          </Button>
        </div>
      </div>

      <ProductPickerModal
        open={pickerOpen}
        onClose={() => setPickerOpen(false)}
        onSelect={handleProductSelect}
      />
    </div>
  );
}
