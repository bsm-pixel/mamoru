'use client';

import { useState } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { CustomerQuickModal } from '@/components/customers/customer-quick-modal';
import { Button } from '@/components/ui/button';
import { formatPhone, formatKRW, formatDate } from '@/lib/utils/format';
import { calcTotalCost } from '@/lib/repair/cost-calculator';
import type { Repair } from '@/lib/supabase/types';
import { User, MapPin, Scissors, Pencil, Check, X } from 'lucide-react';

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

  return (
    <div className="space-y-4">
      {/* 고객 정보 (+ 주소지 통합 — 표기만 한묶음, 저장/송장/연동 로직 무관여) */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>
              <User size={16} className="inline mr-1.5" />
              고객 정보
            </span>
            {onUpdate && !editAddr && (
              <button onClick={() => setEditAddr(true)} className="text-neutral-400 hover:text-neutral-600" title="주소 수정">
                <Pencil size={14} />
              </button>
            )}
          </CardTitle>
        </CardHeader>
        <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
          <dt className="text-neutral-500">이름</dt>
          <dd className="font-medium">
            {r.customer_id ? (
              <CustomerNameLink customerId={r.customer_id} name={r.name} />
            ) : r.name}
          </dd>
          <dt className="text-neutral-500">전화</dt>
          <dd>{formatPhone(r.phone)}</dd>
          <dt className="text-neutral-500">진행방식</dt>
          <dd>
            {r.proceed_type ? (
              <span className="inline-block px-2 py-0.5 rounded-full text-xs font-medium bg-info-soft text-info">
                {r.proceed_type}
              </span>
            ) : '-'}
          </dd>
          <dt className="text-neutral-500">전달방법</dt>
          <dd>{r.delivery_method || '-'}</dd>
          {!editAddr && (
            <>
              <dt className="text-neutral-500">우편번호</dt>
              <dd>{r.postcode || '-'}</dd>
              <dt className="text-neutral-500">주소지</dt>
              <dd>{r.address ? `${r.address}${r.address_detail ? ' ' + r.address_detail : ''}` : '-'}</dd>
            </>
          )}
        </dl>

        {/* 주소 수정 (표기 통합 후에도 편집 기능 동일 유지) */}
        {editAddr && (
          <div className="space-y-2 mt-3 pt-3 border-t border-neutral-100">
            <p className="text-xs text-neutral-500 flex items-center gap-1">
              <MapPin size={12} /> 주소 수정
            </p>
            <input
              type="text"
              value={postcode}
              onChange={(e) => setPostcode(e.target.value)}
              placeholder="우편번호"
              className="w-28 h-8 px-2 rounded border border-neutral-200 text-sm"
            />
            <input
              type="text"
              value={address}
              onChange={(e) => setAddress(e.target.value)}
              placeholder="주소"
              className="w-full h-8 px-2 rounded border border-neutral-200 text-sm"
            />
            <input
              type="text"
              value={addressDetail}
              onChange={(e) => setAddressDetail(e.target.value)}
              placeholder="상세주소"
              className="w-full h-8 px-2 rounded border border-neutral-200 text-sm"
            />
            <div className="flex gap-2 pt-1">
              <Button variant="primary" size="sm" onClick={handleSaveAddr} loading={savingAddr}>
                <Check size={12} /> 저장
              </Button>
              <Button variant="ghost" size="sm" onClick={cancelAddr}>
                <X size={12} /> 취소
              </Button>
            </div>
          </div>
        )}
      </Card>

      {/* 접수 정보 — 수량 수정 가능 */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>
              <Scissors size={16} className="inline mr-1.5" />
              접수 정보
            </span>
            {onUpdate && !editQty && (
              <button onClick={() => setEditQty(true)} className="text-neutral-400 hover:text-neutral-600">
                <Pencil size={14} />
              </button>
            )}
          </CardTitle>
        </CardHeader>
        {editQty ? (
          <div className="space-y-3">
            <div className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
              <dt className="text-neutral-500 pt-1">마모루</dt>
              <dd>
                <input
                  type="number"
                  min={0}
                  value={qtyMamoru}
                  onChange={(e) => setQtyMamoru(parseInt(e.target.value) || 0)}
                  className="w-20 h-8 px-2 rounded border border-neutral-200 text-sm text-right"
                />
                <span className="text-sm ml-1">자루</span>
              </dd>
              <dt className="text-neutral-500 pt-1">타사</dt>
              <dd>
                <input
                  type="number"
                  min={0}
                  value={qtyOther}
                  onChange={(e) => setQtyOther(parseInt(e.target.value) || 0)}
                  className="w-20 h-8 px-2 rounded border border-neutral-200 text-sm text-right"
                />
                <span className="text-sm ml-1">자루</span>
              </dd>
            </div>
            {/* 비용 미리보기 */}
            {previewCost && (
              <div className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-2 space-y-0.5">
                <p>수리비: {formatKRW(previewCost.serviceCost)} / 수거비: {formatKRW(previewCost.shippingFee)}</p>
                <p className="font-semibold text-terracotta-deep">합계: {formatKRW(previewCost.totalAmount)}</p>
              </div>
            )}
            <div className="flex gap-2">
              <Button variant="primary" size="sm" onClick={handleSaveQty} loading={savingQty}>
                <Check size={12} /> 저장
              </Button>
              <Button variant="ghost" size="sm" onClick={cancelQty}>
                <X size={12} /> 취소
              </Button>
            </div>
          </div>
        ) : (
          <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
            <dt className="text-neutral-500">마모루</dt>
            <dd className="font-medium">{r.qty_mamoru}자루</dd>
            <dt className="text-neutral-500">타사</dt>
            <dd className="font-medium">{r.qty_other}자루</dd>
            {r.proceed_type === '직접방문' ? (
              <>
                <dt className="text-neutral-500">방문 예정</dt>
                <dd className="font-medium text-emerald-700">
                  {r.visit_date
                    ? `${formatDate(r.visit_date, 'yyyy.MM.dd')}${r.visit_time ? ` ${String(r.visit_time).slice(0, 5)}` : ''}`
                    : '매장방문 (일시 미정)'}
                </dd>
              </>
            ) : (
              <>
                <dt className="text-neutral-500">수거요청일</dt>
                <dd>{r.pickup_date ? formatDate(r.pickup_date, 'yyyy.MM.dd') : '-'}</dd>
              </>
            )}
            <dt className="text-neutral-500">접수일시</dt>
            <dd>{r.received_at ? formatDate(r.received_at, 'yyyy.MM.dd HH:mm') : '-'}</dd>
          </dl>
        )}
      </Card>

      {/* 비용 정보 — 수동 편집 가능 */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>비용 정보</span>
            {onUpdate && !editCost && (
              <button onClick={() => { setServiceCost(r.service_cost); setShippingFee(r.shipping_fee); setEditCost(true); }} className="text-neutral-400 hover:text-neutral-600">
                <Pencil size={14} />
              </button>
            )}
          </CardTitle>
        </CardHeader>
        {editCost ? (
          <div className="space-y-3">
            <div className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
              <dt className="text-neutral-500 pt-1">수리비</dt>
              <dd>
                <input
                  type="number"
                  min={0}
                  step={1000}
                  value={serviceCost}
                  onChange={(e) => setServiceCost(parseInt(e.target.value) || 0)}
                  className="w-28 h-8 px-2 rounded border border-neutral-200 text-sm text-right"
                />
                <span className="text-xs text-neutral-400 ml-1">원</span>
              </dd>
              <dt className="text-neutral-500 pt-1">수거비</dt>
              <dd>
                <input
                  type="number"
                  min={0}
                  step={1000}
                  value={shippingFee}
                  onChange={(e) => setShippingFee(parseInt(e.target.value) || 0)}
                  className="w-28 h-8 px-2 rounded border border-neutral-200 text-sm text-right"
                />
                <span className="text-xs text-neutral-400 ml-1">원</span>
              </dd>
            </div>
            {/* 합계 미리보기 */}
            <div className="text-xs bg-neutral-50 rounded-lg p-2">
              <span className="text-neutral-500">합계: </span>
              <span className={`font-bold ${serviceCost + shippingFee === 0 ? 'text-green-600' : 'text-terracotta-deep'}`}>
                {serviceCost + shippingFee === 0 ? '무상 처리' : formatKRW(serviceCost + shippingFee)}
              </span>
            </div>
            <div className="flex gap-2">
              <Button variant="primary" size="sm" onClick={handleSaveCost} loading={savingCost}>
                <Check size={12} /> 저장
              </Button>
              <Button variant="ghost" size="sm" onClick={cancelCost}>
                <X size={12} /> 취소
              </Button>
            </div>
          </div>
        ) : (
          <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
            <dt className="text-neutral-500">수리비</dt>
            <dd className="font-medium">{formatKRW(r.service_cost)}</dd>
            <dt className="text-neutral-500">수거비</dt>
            <dd>{formatKRW(r.shipping_fee)}</dd>
            <dt className="text-neutral-500 font-semibold">합계</dt>
            <dd className={`font-bold ${r.total_amount === 0 ? 'text-green-600' : 'text-terracotta-deep'}`}>
              {r.total_amount === 0 ? '무상 처리' : formatKRW(r.total_amount)}
            </dd>
          </dl>
        )}
      </Card>

      {/* 메모 */}
      {r.memo && (
        <Card>
          <CardHeader>
            <CardTitle>고객 메모</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-700 whitespace-pre-wrap">{r.memo}</p>
        </Card>
      )}

      {/* 관리자 메모 */}
      {r.admin_note && (
        <Card>
          <CardHeader>
            <CardTitle>관리자 메모 (추가전달사항)</CardTitle>
          </CardHeader>
          <p className="text-sm text-neutral-700 whitespace-pre-wrap">{r.admin_note}</p>
        </Card>
      )}
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
