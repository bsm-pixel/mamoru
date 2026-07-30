'use client';

import { useState, type ReactNode } from 'react';
import { Card } from '@/components/ui/card';
import { CustomerQuickModal } from '@/components/customers/customer-quick-modal';
import { Button } from '@/components/ui/button';
import { formatPhone, formatKRW, formatDate } from '@/lib/utils/format';
import { calcTotalCost } from '@/lib/repair/cost-calculator';
import type { Repair } from '@/lib/supabase/types';
import { MapPin, Scissors, Pencil, Check, X } from 'lucide-react';
import toast from 'react-hot-toast';

interface RepairDetailCardProps {
  repair: Repair;
  onUpdate?: (fields: Record<string, unknown>) => Promise<void>;
}

export function RepairDetailCard({ repair: r, onUpdate }: RepairDetailCardProps) {
  // 수량 수정 모드
  const [editQty, setEditQty] = useState(false);
  const [qtyMamoru, setQtyMamoru] = useState(r.qty_mamoru);
  const [qtyOther, setQtyOther] = useState(r.qty_other);
  const [savingQty, setSavingQty] = useState(false);

  // 주소 수정 모드
  const [editAddr, setEditAddr] = useState(false);
  const [postcode, setPostcode] = useState(r.postcode || '');
  const [address, setAddress] = useState(r.address || '');
  const [addressDetail, setAddressDetail] = useState(r.address_detail || '');
  const [savingAddr, setSavingAddr] = useState(false);

  // 비용 수동 편집 모드
  const [editCost, setEditCost] = useState(false);
  const [serviceCost, setServiceCost] = useState(r.service_cost);
  const [shippingFee, setShippingFee] = useState(r.shipping_fee);
  const [savingCost, setSavingCost] = useState(false);

  const handleSaveQty = async () => {
    if (!onUpdate) return;
    setSavingQty(true);
    try {
      const costs = calcTotalCost(qtyMamoru, qtyOther, r.proceed_type);
      await onUpdate({
        qty_mamoru: qtyMamoru,
        qty_other: qtyOther,
        service_cost: costs.serviceCost,
        shipping_fee: costs.shippingFee,
        total_amount: costs.totalAmount,
      });
      setEditQty(false);
    } finally {
      setSavingQty(false);
    }
  };

  const handleSaveAddr = async () => {
    if (!onUpdate) return;
    setSavingAddr(true);
    try {
      await onUpdate({ postcode, address, address_detail: addressDetail });
      setEditAddr(false);
    } finally {
      setSavingAddr(false);
    }
  };

  const cancelQty = () => {
    setQtyMamoru(r.qty_mamoru);
    setQtyOther(r.qty_other);
    setEditQty(false);
  };

  const cancelAddr = () => {
    setPostcode(r.postcode || '');
    setAddress(r.address || '');
    setAddressDetail(r.address_detail || '');
    setEditAddr(false);
  };

  const handleSaveCost = async () => {
    if (!onUpdate) return;
    setSavingCost(true);
    try {
      await onUpdate({
        service_cost: serviceCost,
        shipping_fee: shippingFee,
        total_amount: serviceCost + shippingFee,
      });
      setEditCost(false);
    } finally {
      setSavingCost(false);
    }
  };

  const cancelCost = () => {
    setServiceCost(r.service_cost);
    setShippingFee(r.shipping_fee);
    setEditCost(false);
  };

  // 수량 변경 시 비용 미리보기
  const previewCost = editQty ? calcTotalCost(qtyMamoru, qtyOther, r.proceed_type) : null;

  const visitOrPickup = r.proceed_type === '직접방문'
    ? { label: '방문 예정', value: r.visit_date ? `${formatDate(r.visit_date, 'yyyy.MM.dd')}${r.visit_time ? ` ${String(r.visit_time).slice(0, 5)}` : ''}` : '매장방문 (일시 미정)', cls: 'text-emerald-700 font-medium' }
    : { label: '수거요청일', value: r.pickup_date ? formatDate(r.pickup_date, 'yyyy.MM.dd') : '-', cls: '' };
  const addrText = r.address ? `${r.address}${r.address_detail ? ' ' + r.address_detail : ''}${r.postcode ? ` (${r.postcode})` : ''}` : '';
  const anyEdit = editQty || editAddr || editCost;

  return (
    // PC-네이티브: 정보 카드 1개로 압축 + 2열 그리드 (판매 상세 패널 톤). 편집은 인라인 확장.
    <Card>
      <div className="@container space-y-3">
        {/* 상단: 고객명 + 진행방식 칩 + 편집 툴바 */}
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <Scissors size={15} className="text-neutral-400 shrink-0" />
            <span className="font-bold text-sm truncate">
              {r.customer_id ? <CustomerNameLink customerId={r.customer_id} name={r.name} /> : r.name}
            </span>
            {r.proceed_type && (
              <span className="shrink-0 inline-block px-2 py-0.5 rounded-full text-[11px] font-medium bg-info-soft text-info">{r.proceed_type}</span>
            )}
          </div>
          {onUpdate && !anyEdit && (
            <div className="flex items-center gap-1 shrink-0 text-neutral-400">
              <button onClick={() => setEditQty(true)} title="수량 수정" className="px-1.5 py-0.5 rounded hover:bg-neutral-100 text-[11px] font-medium">수량<Pencil size={10} className="inline ml-0.5" /></button>
              <button onClick={() => setEditAddr(true)} title="주소 수정" className="px-1.5 py-0.5 rounded hover:bg-neutral-100 text-[11px] font-medium">주소<Pencil size={10} className="inline ml-0.5" /></button>
            </div>
          )}
        </div>

        {/* 정보 — 2열 그리드 (라벨 위 / 값 아래, 판매 패널 방식) */}
        <div className="grid grid-cols-2 gap-x-4 gap-y-2.5 text-sm">
          <InfoField label="전화">{formatPhone(r.phone)}</InfoField>
          <InfoField label="전달방법">{r.delivery_method || '-'}</InfoField>
          <InfoField label="마모루">{r.qty_mamoru}자루</InfoField>
          <InfoField label="타사">{r.qty_other}자루</InfoField>
          <InfoField label={visitOrPickup.label} valueClass={visitOrPickup.cls}>{visitOrPickup.value}</InfoField>
          <InfoField label="접수일시">{r.received_at ? formatDate(r.received_at, 'yyyy.MM.dd HH:mm') : '-'}</InfoField>
          {addrText && <InfoField label="주소지" full>{addrText}</InfoField>}
        </div>

        {/* 비용 요약 — 한 줄 (수리비·수거비 → 합계 + 편집) */}
        <div className="flex items-center justify-between gap-2 border-t border-neutral-100 pt-2.5">
          <span className="text-xs text-neutral-500">수리비 {formatKRW(r.service_cost)} · 수거비 {formatKRW(r.shipping_fee)}</span>
          <span className="flex items-center gap-1.5">
            <span className={`text-sm font-bold ${r.total_amount === 0 ? 'text-green-600' : 'text-terracotta-deep'}`}>
              {r.total_amount === 0 ? '무상' : formatKRW(r.total_amount)}
            </span>
            {onUpdate && !anyEdit && (
              <button onClick={() => { setServiceCost(r.service_cost); setShippingFee(r.shipping_fee); setEditCost(true); }} title="비용 수정" className="text-neutral-400 hover:text-neutral-600">
                <Pencil size={12} />
              </button>
            )}
          </span>
        </div>

        {/* 121: 직접방문 고객 셀프 링크 (컴팩트 텍스트 링크) */}
        {r.proceed_type === '직접방문' && (r as { manage_token?: string | null }).manage_token && (
          <button
            onClick={() => {
              const link = `https://page.mamoru.kr/projects/as/page_change_request.html?uid=${(r as { manage_token?: string | null }).manage_token}`;
              navigator.clipboard?.writeText(link);
              toast.success('고객 관리 링크 복사됨');
            }}
            className="text-[11px] font-medium text-neutral-500 hover:text-blue-600 transition"
          >
            🔗 고객 일정변경/취소 링크 복사
          </button>
        )}

        {/* 메모 (있을 때만, 컴팩트) */}
        {(r.memo || r.admin_note) && (
          <div className="border-t border-neutral-100 pt-2.5 space-y-1.5 text-sm">
            {r.memo && <p className="text-neutral-700 whitespace-pre-wrap"><span className="text-xs text-neutral-400 mr-1">고객메모</span>{r.memo}</p>}
            {r.admin_note && <p className="text-neutral-700 whitespace-pre-wrap"><span className="text-xs text-neutral-400 mr-1">관리자</span>{r.admin_note}</p>}
          </div>
        )}

        {/* ── 편집 모드 (인라인 확장) ── */}
        {editQty && (
          <div className="border-t border-neutral-100 pt-3 space-y-2">
            <p className="text-xs font-semibold text-neutral-500">수량 수정</p>
            <div className="flex items-center gap-3 text-sm">
              <label className="flex items-center gap-1.5">마모루
                <input type="number" min={0} value={qtyMamoru} onChange={(e) => setQtyMamoru(parseInt(e.target.value) || 0)} className="w-16 h-8 px-2 rounded border border-neutral-200 text-right" />자루
              </label>
              <label className="flex items-center gap-1.5">타사
                <input type="number" min={0} value={qtyOther} onChange={(e) => setQtyOther(parseInt(e.target.value) || 0)} className="w-16 h-8 px-2 rounded border border-neutral-200 text-right" />자루
              </label>
            </div>
            {previewCost && (
              <div className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-2">
                수리비 {formatKRW(previewCost.serviceCost)} · 수거비 {formatKRW(previewCost.shippingFee)} → <span className="font-semibold text-terracotta-deep">합계 {formatKRW(previewCost.totalAmount)}</span>
              </div>
            )}
            <div className="flex gap-2">
              <Button variant="primary" size="sm" onClick={handleSaveQty} loading={savingQty}><Check size={12} /> 저장</Button>
              <Button variant="ghost" size="sm" onClick={cancelQty}><X size={12} /> 취소</Button>
            </div>
          </div>
        )}
        {editCost && (
          <div className="border-t border-neutral-100 pt-3 space-y-2">
            <p className="text-xs font-semibold text-neutral-500">비용 수정</p>
            <div className="flex items-center gap-3 text-sm">
              <label className="flex items-center gap-1.5">수리비
                <input type="number" min={0} step={1000} value={serviceCost} onChange={(e) => setServiceCost(parseInt(e.target.value) || 0)} className="w-24 h-8 px-2 rounded border border-neutral-200 text-right" />원
              </label>
              <label className="flex items-center gap-1.5">수거비
                <input type="number" min={0} step={1000} value={shippingFee} onChange={(e) => setShippingFee(parseInt(e.target.value) || 0)} className="w-24 h-8 px-2 rounded border border-neutral-200 text-right" />원
              </label>
            </div>
            <div className="text-xs bg-neutral-50 rounded-lg p-2">
              합계 <span className={`font-bold ${serviceCost + shippingFee === 0 ? 'text-green-600' : 'text-terracotta-deep'}`}>{serviceCost + shippingFee === 0 ? '무상 처리' : formatKRW(serviceCost + shippingFee)}</span>
            </div>
            <div className="flex gap-2">
              <Button variant="primary" size="sm" onClick={handleSaveCost} loading={savingCost}><Check size={12} /> 저장</Button>
              <Button variant="ghost" size="sm" onClick={cancelCost}><X size={12} /> 취소</Button>
            </div>
          </div>
        )}
        {editAddr && (
          <div className="border-t border-neutral-100 pt-3 space-y-2">
            <p className="text-xs font-semibold text-neutral-500 flex items-center gap-1"><MapPin size={12} /> 주소 수정</p>
            <input type="text" value={postcode} onChange={(e) => setPostcode(e.target.value)} placeholder="우편번호" className="w-28 h-8 px-2 rounded border border-neutral-200 text-sm" />
            <input type="text" value={address} onChange={(e) => setAddress(e.target.value)} placeholder="주소" className="w-full h-8 px-2 rounded border border-neutral-200 text-sm" />
            <input type="text" value={addressDetail} onChange={(e) => setAddressDetail(e.target.value)} placeholder="상세주소" className="w-full h-8 px-2 rounded border border-neutral-200 text-sm" />
            <div className="flex gap-2">
              <Button variant="primary" size="sm" onClick={handleSaveAddr} loading={savingAddr}><Check size={12} /> 저장</Button>
              <Button variant="ghost" size="sm" onClick={cancelAddr}><X size={12} /> 취소</Button>
            </div>
          </div>
        )}
      </div>
    </Card>
  );
}

/** 정보 필드 — 라벨 위(작게) / 값 아래 (판매 패널 그리드 방식) */
function InfoField({ label, children, full, valueClass }: { label: string; children: ReactNode; full?: boolean; valueClass?: string }) {
  return (
    <div className={full ? 'col-span-2' : ''}>
      <span className="text-xs text-neutral-500">{label}</span>
      <p className={valueClass || 'font-medium'}>{children}</p>
    </div>
  );
}

/** 고객명 클릭 → 퀵뷰 모달 */
function CustomerNameLink({ customerId, name }: { customerId: string; name: string }) {
  const [open, setOpen] = useState(false);
  return (
    <>
      <button onClick={() => setOpen(true)} className="text-blue-600 hover:underline">{name}</button>
      <CustomerQuickModal customerId={customerId} open={open} onClose={() => setOpen(false)} />
    </>
  );
}
