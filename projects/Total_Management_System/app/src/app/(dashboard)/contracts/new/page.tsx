'use client';

import { MamoruWordmark } from '@/components/ui/mamoru-logo';
import { useUnsavedGuard } from '@/hooks/use-unsaved-guard';
import { useFormDraft, DraftBanner } from '@/hooks/use-form-draft';
import { useState, useMemo, useRef } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { SignPadField } from '@/components/contracts/sign-pad';
import { DaumPostcodeButton } from '@/components/shared/daum-postcode-button';
import { typingPhone } from '@/lib/utils/validate';
import { ProductPickerModal } from '@/components/contracts/product-picker-modal';
import { TodayConsultationPicker } from '@/components/contracts/today-consultation-picker';
import { useCreateContract } from '@/hooks/use-contracts';
import { formatKRW } from '@/lib/utils/format';
import { Plus, X, Users } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';
import { MissingNote } from '@/components/ui/missing-note';
import { CAUTION_NOTICE, LEGAL_NOTICE, ENGRAVING_NOTICE } from '@/lib/contracts/notice';
import { MoneyInput } from '@/components/ui/money-input';
import toast from 'react-hot-toast';

/* -- 제품 행 -- */
interface ProductRow {
  id: string;
  product: Product | null;
  quantity: number;
}
export default function NewContractPage() {
  const router = useRouter();
  const createContract = useCreateContract();
  const contractRef = useRef<HTMLDivElement>(null);

  /* -- 고객 정보 (DB 저장용 텍스트) -- */
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  // 2026-11-23 주소 = 주소 검색(도로명) + 상세 주소. 계약서에는 합쳐서 저장, 고객 등록에는 나눠서 넘긴다
  const [addrRoad, setAddrRoad] = useState('');
  const [addrDetail, setAddrDetail] = useState('');
  const [postcode, setPostcode] = useState('');
  const customerAddress = [addrRoad, addrDetail].map((x) => x.trim()).filter(Boolean).join(' ');
  const [customerId, setCustomerId] = useState<string | null>(null);
  const [consultationId, setConsultationId] = useState<string | null>(null);

  /* -- 매장 정보 -- */
  const [shopName, setShopName] = useState('');

  /* -- 상담자 불러오기 모달 -- */
  const [consultPickerOpen, setConsultPickerOpen] = useState(false);

  /* -- 제품 -- */
  const [rows, setRows] = useState<ProductRow[]>([{ id: crypto.randomUUID(), product: null, quantity: 1 }]);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [activeRowId, setActiveRowId] = useState<string | null>(null);

  /* -- 결제 -- */
  const [paymentMethod, setPaymentMethod] = useState('transfer');
  const [installment, setInstallment] = useState(0);
  const [discount, setDiscount] = useState(0);
  const [depositAmount, setDepositAmount] = useState(0);
  const [memo, setMemo] = useState('');

  /* -- 서명 -- */
  const [buyerSignature, setBuyerSignature] = useState('');
  const [sellerSignature, setSellerSignature] = useState('');
  // 2026-11-23 서명 전에 「위 내용을 확인했습니다」 체크
  const [confirmed, setConfirmed] = useState(false);
  // 2026-11-22 서명 받는 방법: 이 자리(매장·출장 화면) / 고객 폰으로 링크 (온라인·톡상담 고객)
  const [signMode, setSignMode] = useState<'here' | 'link'>('here');

  /* -- 이니셜 각인 (2026-11-22) — 문구는 지금 적거나, 나중에 받는다(고객이 서명 링크로 보내거나 직원이 상세에서 적는다) -- */
  const [engraving, setEngraving] = useState(false);
  const [engravingLater, setEngravingLater] = useState(false);
  const [engravingText, setEngravingText] = useState('');

  /* -- 계산 -- */
  const totalAmount = useMemo(() =>
    rows.reduce((s, r) => s + (r.product?.price || 0) * r.quantity, 0), [rows]);
  const finalAmount = totalAmount - discount;
  const balanceAmount = Math.max(0, finalAmount - depositAmount);

  /* -- 날짜 -- */
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');

  /* -- 상담자 불러오기 -- */
  function handleConsultationSelect(consultation: {
    id: string;
    customer_id: string | null;
    name: string;
    phone: string;
    address_road: string | null;
    address_detail: string | null;
  }) {
    setCustomerName(consultation.name);
    setCustomerPhone(typingPhone(consultation.phone || ''));
    setAddrRoad(consultation.address_road || '');
    setAddrDetail(consultation.address_detail || '');
    setPostcode('');
    setCustomerId(consultation.customer_id);
    setConsultationId(consultation.id);
    setConsultPickerOpen(false);
  }

  /* -- 제품 모달 -- */
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

  /* -- 제출 -- */
  async function handleSubmit() {
    const filledRows = rows.filter((r) => r.product);
    if (filledRows.length === 0) return;
    if (signMode === 'link' && !customerName.trim()) { toast.error('링크로 보낼 고객 이름을 적어 주세요'); return; }
    const name = customerName.trim() || '미입력';
    const byLink = signMode === 'link';
    const engText = engraving && !engravingLater ? engravingText.trim() : '';

    const result = await createContract.mutateAsync({
      contract: {
        customer_id: customerId || undefined,
        customer_name: name,
        customer_phone: customerPhone.trim() || undefined,
        customer_address: customerAddress.trim() || undefined,
        total_amount: totalAmount,
        discount_amount: discount,
        final_amount: finalAmount,
        payment_method: paymentMethod,
        installment_months: installment,
        signature_data: (!byLink && buyerSignature) || undefined, // 링크로 보낼 때는 고객이 폰에서 서명
        terms_agreed: !byLink && !!buyerSignature && confirmed,
        memo: memo.trim() || undefined,
        engraving,
        engraving_text: engText || undefined,
        delivery_method: 'shipping',
        deposit_amount: depositAmount,
        balance_amount: balanceAmount,
        seller_signature: sellerSignature || undefined,
        shop_name: shopName.trim() || undefined,
        shop_address: customerAddress.trim() || undefined,
        consultation_id: consultationId || undefined,
        // 2026-11-23 성함·연락처·주소는 손글씨 대신 입력 칸 — handwriting_* 는 더 보내지 않는다(옛 계약서는 그대로 보임)
        customer_address_road: addrRoad.trim() || undefined,
        customer_address_detail: addrDetail.trim() || undefined,
        customer_postcode: postcode || undefined,
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

    // 링크로 보내기: 서명 링크를 만들고 상세로 — 상세의 「고객 서명 링크」 칸에서 톡 문구를 복사한다
    if (byLink && result?.contract?.id) {
      try {
        const r = await fetch(`/api/contracts/${result.contract.id}/sign-link`, { method: 'POST' });
        if (!r.ok) throw new Error(await r.text());
      } catch {
        toast.error('서명 링크를 만들지 못했어요 — 계약서 상세에서 [링크 만들기]를 다시 눌러 주세요');
      }
      draft.clear();
      router.push(`/contracts/${result.contract.id}`);
      return;
    }

    // 계약서 이미지 캡처 + Storage 업로드 (실패해도 저장은 완료)
    if (contractRef.current && result?.contract?.id) {
      try {
        const html2canvas = (await import('html2canvas')).default;
        const canvas = await html2canvas(contractRef.current, {
          scale: 2,
          useCORS: true,
          backgroundColor: '#ffffff',
        });
        const imageBase64 = canvas.toDataURL('image/png');
        await fetch(`/api/contracts/${result.contract.id}/image`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ imageBase64 }),
        });
      } catch (e) {
        console.error('[contract image capture] failed:', e);
      }
    }

    draft.clear(); // 저장 성공 — 임시 저장분을 지운다
    // 2026-11-23 상세로 — 서명된 계약서는 거기서 [판매 입력] 한 번(고객·제품이 채워진 채로)
    router.push(result?.contract?.id ? `/contracts/${result.contract.id}` : '/contracts');
  }

  const canSubmit = rows.some((r) => r.product) && !createContract.isPending && (signMode === 'here' || !!customerName.trim());

  /* -- 입력 스타일 -- */
  // 2026-11-23 휴대폰 글자 16px(아이폰 자동 확대 방지) · 누름 높이 44px
  const inputClass = 'border-0 border-b border-neutral-300 bg-transparent px-1 py-2 min-h-11 text-base sm:text-sm focus:outline-none focus:border-neutral-800 w-full';

  // 2026-11-15 넣다가 다른 메뉴를 누르면 한 번 묻는다 (저장 성공 뒤 스스로 이동할 때는 묻지 않음)
  const unsavedGuard = useUnsavedGuard(!!customerName.trim() || rows.some((r) => !!r.product) || !!memo.trim() || !!buyerSignature);

  // 2026-11-21 임시 저장 — 서명과 손글씨 이미지는 넣지 않는다(크고, 함께 쓰는 태블릿에 남으면 안 된다). 이어서 쓸 때 서명은 다시 받는다
  // 2026-11-22 판 2: 서명 받는 방법 · 각인 칸 추가 (판 1 임시 저장분은 버린다)
  // 2026-11-23 판 3: 주소를 도로명·상세로 나눔 (판 2 임시 저장분은 버린다)
  const draft = useFormDraft('contract-new', 3, { customerName, customerPhone, addrRoad, addrDetail, postcode, customerId, consultationId, shopName, rows, paymentMethod, installment, discount, depositAmount, memo, signMode, engraving, engravingLater, engravingText }, !!customerName.trim() || rows.some((r) => !!r.product) || !!memo.trim());
  const resumeDraft = () => {
    const d = draft.take();
    if (!d) return;
    setCustomerName(d.customerName); setCustomerPhone(d.customerPhone); setAddrRoad(d.addrRoad); setAddrDetail(d.addrDetail); setPostcode(d.postcode); setCustomerId(d.customerId); setConsultationId(d.consultationId);
    setShopName(d.shopName); setRows(d.rows); setPaymentMethod(d.paymentMethod); setInstallment(d.installment); setDiscount(d.discount); setDepositAmount(d.depositAmount); setMemo(d.memo);
    setSignMode(d.signMode); setEngraving(d.engraving); setEngravingLater(d.engravingLater); setEngravingText(d.engravingText);
  };

  return (
    <>
      {unsavedGuard}
      <Topbar title="계약서 작성" />

      <div className="px-3 py-4">
        <div className="max-w-[600px] mx-auto mb-2"><DraftBanner at={draft.pendingAt} onResume={resumeDraft} onDiscard={draft.discard} /></div>
        {/* 2026-11-22 서명 받는 방법 — 계약서 종이 밖(캡처에 안 들어감) */}
        <div className="max-w-[600px] mx-auto mb-3" data-sign-mode>
          <p className="text-xs font-semibold text-neutral-600 mb-1.5">서명 받는 방법</p>
          <div className="grid grid-cols-2 gap-1 p-1 rounded-lg bg-neutral-100" role="radiogroup" aria-label="서명 받는 방법">
            {([['here', '이 자리에서 서명'], ['link', '고객에게 링크로 보내기']] as const).map(([v, label]) => (
              <button key={v} type="button" role="radio" aria-checked={signMode === v} data-mode={v}
                onClick={() => setSignMode(v)}
                className={`min-h-10 rounded-md text-sm font-semibold transition ${signMode === v ? 'bg-white text-neutral-900 shadow-sm' : 'text-neutral-500 hover:text-neutral-800'}`}>
                {label}
              </button>
            ))}
          </div>
          {signMode === 'link' && (
            <p className="text-xs text-neutral-500 mt-1.5">저장하면 서명 링크가 만들어져요. 톡으로 보내면 고객이 폰에서 내용을 읽고 서명합니다.</p>
          )}
        </div>
        <div ref={contractRef} className="bg-white max-w-[600px] mx-auto border border-neutral-400 rounded-sm shadow-sm">

          {/* -- 헤더 -- */}
          <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-400">
            <MamoruWordmark height={10} color="#262626" />
            <h2 className="text-base font-extrabold tracking-[0.3em] text-neutral-900">구 매 계 약 서</h2>
          </div>

          {/* -- 상담자 불러오기 + 매장명 -- */}
          <div className="px-5 py-4 border-b border-neutral-300 space-y-3">
            <div className="flex items-center gap-2">
              <button
                onClick={() => setConsultPickerOpen(true)}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-neutral-100 border border-neutral-200 text-xs font-semibold text-neutral-700 hover:bg-neutral-200 transition"
              >
                <Users size={14} />
                상담자 불러오기
              </button>
              {consultationId && (
                <span className="text-[10px] text-green-600 font-medium">상담 연결됨</span>
              )}
            </div>
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
          </div>

          {/* -- 고객 정보 (2026-11-23 손글씨 칸 → 키보드 입력 칸. 상담자 불러오기로 채워지면 고객은 확인만) -- */}
          <div className="px-5 py-4 border-b border-neutral-300 space-y-3" data-link-customer data-customer-fields>
            <div>
              <label htmlFor="ct-name" className="text-[11px] text-neutral-500">성함</label>
              <input id="ct-name" type="text" autoComplete="name" value={customerName} onChange={(e) => setCustomerName(e.target.value)} placeholder="성함" className={inputClass} />
            </div>
            <div>
              <label htmlFor="ct-phone" className="text-[11px] text-neutral-500">연락처</label>
              <input id="ct-phone" type="tel" inputMode="numeric" autoComplete="tel" value={customerPhone} onChange={(e) => setCustomerPhone(typingPhone(e.target.value))} placeholder="010-0000-0000" className={inputClass} />
            </div>
            <div>
              <label htmlFor="ct-addr" className="text-[11px] text-neutral-500">주소 (수령지)</label>
              <div className="flex items-end gap-2">
                <input id="ct-addr" type="text" value={addrRoad} onChange={(e) => setAddrRoad(e.target.value)} placeholder="주소 검색을 눌러 주세요" className={inputClass} />
                <DaumPostcodeButton onSelected={(d) => { setAddrRoad(d.roadAddress); setPostcode(d.zonecode); document.getElementById('ct-addr-detail')?.focus(); }}
                  className="shrink-0 min-h-11 px-3 rounded-md border border-neutral-800 text-xs font-semibold text-neutral-800 hover:bg-neutral-100 transition">
                  주소 검색
                </DaumPostcodeButton>
              </div>
              <input id="ct-addr-detail" type="text" aria-label="상세 주소" value={addrDetail} onChange={(e) => setAddrDetail(e.target.value)} placeholder="상세 주소 (동·호수)" className={inputClass + ' mt-1'} />
            </div>
          </div>

          {/* -- 필독 (유의사항) -- */}
          <div className="px-5 py-4 border-b border-neutral-300 bg-neutral-50">
            <p className="text-xs font-bold text-neutral-700 mb-2 text-center">※ 필독 ※</p>
            <p className="text-[11px] text-neutral-600 leading-relaxed whitespace-pre-line">{LEGAL_NOTICE}</p>
            <div className="mt-3 pt-2 border-t border-neutral-200">
              <p className="text-[10px] font-bold text-neutral-700 mb-1">[유의사항]</p>
              <p className="text-[11px] text-neutral-600 leading-relaxed whitespace-pre-line">{CAUTION_NOTICE}</p>
            </div>
          </div>

          {/* -- 결제방식 -- */}
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
                <MoneyInput
                  value={depositAmount}
                  onChange={(val) => setDepositAmount(val)}
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
              <MoneyInput
                value={discount}
                onChange={(val) => setDiscount(val)}
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

          {/* -- 제품 테이블 -- */}
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
                        <button title="삭제" aria-label="삭제" onClick={() => removeRow(row.id)} className="text-neutral-400 hover:text-red-500">
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
              className="flex items-center gap-1.5 mt-3 text-xs text-neutral-500 hover:text-neutral-800 transition"
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

          {/* -- 이니셜 각인 (2026-11-22) — 체크했을 때만 안내 문구가 보인다 -- */}
          <div className="px-5 py-4 border-b border-neutral-300 space-y-3" data-engraving>
            <label className="flex items-center gap-2 text-sm font-semibold text-neutral-800 cursor-pointer min-h-8">
              <input type="checkbox" checked={engraving} onChange={(e) => setEngraving(e.target.checked)} className="w-4 h-4 accent-neutral-800" data-engraving-check />
              이니셜 각인
            </label>
            {engraving && (
              <>
                <p className="text-[11px] text-neutral-700 leading-relaxed border-l-2 border-neutral-800 pl-3 py-1" data-engraving-notice>{ENGRAVING_NOTICE}</p>
                <div className="flex flex-wrap gap-4 text-sm" role="radiogroup" aria-label="각인 문구">
                  <label className="flex items-center gap-1.5 cursor-pointer">
                    <input type="radio" name="engraving-when" checked={!engravingLater} onChange={() => setEngravingLater(false)} className="accent-neutral-800" />
                    지금 적기
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer">
                    <input type="radio" name="engraving-when" checked={engravingLater} onChange={() => setEngravingLater(true)} className="accent-neutral-800" data-engraving-later />
                    문구는 나중에 받기
                  </label>
                </div>
                {!engravingLater ? (
                  <div>
                    <label className="text-[10px] text-neutral-500">각인 문구</label>
                    <input type="text" value={engravingText} onChange={(e) => setEngravingText(e.target.value)} maxLength={30} placeholder="예: S.M.B" className={inputClass} />
                  </div>
                ) : (
                  <p className="text-[11px] text-neutral-500 leading-relaxed">
                    {signMode === 'link'
                      ? '고객이 서명 링크에서 바로 적거나, 정해지면 같은 링크로 보낼 수 있어요.'
                      : '문구가 정해지면 서명 링크로 받거나, 톡으로 받아 계약서 상세에서 적어 넣어요.'}
                  </p>
                )}
              </>
            )}
          </div>

          {/* -- 날짜 + 서명 -- */}
          <div className="px-5 py-4 border-b border-neutral-300 space-y-4">
            <p className="text-center text-sm text-neutral-800 tracking-widest">
              {yyyy} 년 {mm} 월 {dd} 일
            </p>

            <div className="space-y-4">
              {signMode === 'link' ? (
                <p className="text-xs text-neutral-500 border border-dashed border-neutral-300 rounded-md px-3 py-2">구매자 서명은 고객이 링크에서 합니다</p>
              ) : (
                <div className="space-y-2" data-buyer-sign>
                  {/* 2026-11-23 손으로 쓰는 것은 이 최종 서명 하나 — 확인 체크 → [서명하기] → 화면 전체 서명판 */}
                  <label className="flex items-center gap-2 text-sm font-semibold text-neutral-800 cursor-pointer min-h-11">
                    <input type="checkbox" checked={confirmed} onChange={(e) => setConfirmed(e.target.checked)} className="w-5 h-5 accent-neutral-800" data-confirm-check />
                    위 내용을 확인했습니다
                  </label>
                  <SignPadField value={buyerSignature} onChange={setBuyerSignature} label="구매자 서명" testId="buyer"
                    lockedReason={confirmed ? null : '위 내용을 확인하고 체크하면 서명할 수 있어요'} />
                </div>
              )}
              <div className="space-y-2">
                <p className="text-xs font-bold text-neutral-700">판매자 서명</p>
                <SignPadField value={sellerSignature} onChange={setSellerSignature} label="판매자 서명" testId="seller" />
              </div>
            </div>
          </div>

          {/* -- 계좌 안내 -- */}
          <div className="px-5 py-3 bg-neutral-50 text-center">
            <p className="text-[11px] text-neutral-600">
              입금 계좌: <span className="font-semibold">우리은행 1002-439-462514 (백성민)</span>
            </p>
          </div>
        </div>

        {/* -- 저장 버튼 -- */}
        <div className="max-w-[600px] mx-auto mt-4 pb-8">
          <MissingNote className="mb-1.5 text-center" items={[!rows.some((r) => r.product) && '제품 1개 이상', signMode === 'link' && !customerName.trim() && '고객 이름']} />
          <Button
            className="w-full"
            disabled={!canSubmit}
            onClick={handleSubmit}
            data-contract-save
          >
            {createContract.isPending ? '저장 중...' : signMode === 'link' ? `저장하고 서명 링크 만들기 (${formatKRW(finalAmount)})` : `계약서 저장 (${formatKRW(finalAmount)})`}
          </Button>
        </div>
      </div>

      <ProductPickerModal
        open={pickerOpen}
        onClose={() => setPickerOpen(false)}
        onSelect={handleProductSelect}
      />

      <TodayConsultationPicker
        open={consultPickerOpen}
        onClose={() => setConsultPickerOpen(false)}
        onSelect={handleConsultationSelect}
      />
    </>
  );
}
