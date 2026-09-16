'use client';

import { useState } from 'react';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { useDelivery, useUpdateDelivery } from '@/hooks/use-deliveries';
import { useProducts } from '@/hooks/use-sales';
import { useCustomerCatalog } from '@/hooks/use-customer-catalog';
import type { Product } from '@/lib/supabase/types';
import { formatKRW, formatDate, calcVAT, CUSTOMER_TYPE_LABEL, CUSTOMER_TYPE_COLOR } from '@/lib/utils/format';
import { useQueryClient } from '@tanstack/react-query';
import { Package, Pencil, Save, X, Printer, Minus, Plus, Trash2 } from 'lucide-react';
import toast from 'react-hot-toast';
import { DLPrintModal } from './dl-print-modal';
import { getDeliveryStatusChip, isAwaitingPickup } from '@/lib/deliveries/status';
import { COURIER_OPTIONS, COURIER_LOTTE, COURIER_DIRECT, isAlpsTrackable, isDirectHandover, courierLabel } from '@/lib/shipping/couriers';
import { ActionNote, MoreActions, DangerZone, DangerLink, SubtleButton } from '@/components/ui/action-section';
import { useEnterSelect } from '@/hooks/use-enter-select';

// 상태 뱃지 라벨/색은 lib/deliveries/status.ts getDeliveryStatusChip 이 단일 출처 (여기 있던 사본은 미사용이라 제거)
const PAYMENT_LABEL: Record<string, string> = { unpaid: '미결제', partial: '부분결제', paid: '결제완료' };
const PAYMENT_COLOR: Record<string, string> = { unpaid: 'bg-red-100 text-red-600', partial: 'bg-yellow-100 text-yellow-700', paid: 'bg-green-100 text-green-700' };
const RECEIPT_LABEL: Record<string, string> = { expense_proof: '지출증빙', tax_invoice: '세금계산서', none: '미적용' };
// 고객유형 라벨/색은 format.ts SSOT 사용

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type DL = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type DLItem = any;

interface Props {
  deliveryId: string;
}

export function DeliveryDetailPanel({ deliveryId }: Props) {
  const { data, isLoading } = useDelivery(deliveryId);
  const updateDL = useUpdateDelivery();
  const queryClient = useQueryClient();
  const { data: products = [] } = useProducts();
  // 거래처별 납품명·가격 (생성 모달과 동일 — catalog 우선)
  const { data: customerCatalogData } = useCustomerCatalog((data?.delivery?.customer_id as string | undefined) ?? undefined);

  const [editing, setEditing] = useState(false);
  const [editItems, setEditItems] = useState<Array<{
    product_id?: string; product_name: string; sku?: string; category?: string; quantity: number; unit_price: number;
  }>>([]);
  const [editMemo, setEditMemo] = useState('');
  const [addSearch, setAddSearch] = useState('');
  const [editExpectedDate, setEditExpectedDate] = useState('');
  const [saving, setSaving] = useState(false);
  const [showPrint, setShowPrint] = useState(false);
  const [trackingInput, setTrackingInput] = useState('');
  // 150: 수동 출고 시 택배사 — 롯데 외 배송·직접전달 경로.
  //      null = 아직 안 고름 → 이 건에 이미 기록된 택배사를 기본값으로 쓴다(우체국 건을 열면 우체국이 뜨게)
  const [courierPick, setCourierPick] = useState<string | null>(null);
  const [bookingInvoice, setBookingInvoice] = useState(false);
  const [pendingAction, setPendingAction] = useState<{
    action: string; label: string; msg: string; variant?: 'danger' | 'default'; extra?: Record<string, unknown>;
  } | null>(null);

  // 제품 추가 후보 — 이미 담은 것 제외 + 이름·SKU 검색 (훅이라 early return 위에 둔다)
  const addableProducts = products.filter((p) => {
    if (editItems.find((ei) => ei.product_id === p.id)) return false;
    const q = addSearch.trim().toLowerCase();
    if (!q) return true;
    return p.name.toLowerCase().includes(q) || (p.sku || '').toLowerCase().includes(q);
  });
  // 검색 결과에서 ↑↓·Enter 로 바로 담기 (결과 1개면 Enter 만으로)
  const addPick = useEnterSelect(addableProducts, (p: Product) => { addEditItem(p); setAddSearch(''); });

  if (isLoading) {
    return <div className="p-4 space-y-4"><Skeleton className="h-48" /><Skeleton className="h-32" /></div>;
  }
  if (!data?.delivery) {
    return (
      <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
        <Package size={28} className="mb-2 opacity-40" />
        <p className="text-sm">납품 정보를 찾을 수 없습니다</p>
      </div>
    );
  }

  const dl: DL = data.delivery;
  const items: DLItem[] = data.items;
  const status = (dl.status as string) || 'draft';
  const courierInput = courierPick ?? ((dl.courier_name as string | null) || COURIER_LOTTE);

  // 거래처별 납품명·가격 결정 (생성 모달 create-delivery-modal 과 동일 규칙: catalog → 고객유형 → 기본)
  const catalogEntryMap = new Map((customerCatalogData?.catalog || []).map((c) => [c.product_id, c]));
  const dlCustomerType = dl.customer_type as string | undefined;
  function getDeliveryPrice(p: Product): number {
    const ce = catalogEntryMap.get(p.id);
    if (ce?.unit_price && ce.unit_price > 0) return ce.unit_price;
    if (dlCustomerType === 'dealer' && (p as Record<string, unknown>).price_dealer) return (p as Record<string, unknown>).price_dealer as number;
    if (dlCustomerType === 'academy' && (p as Record<string, unknown>).price_academy) return (p as Record<string, unknown>).price_academy as number;
    return p.price;
  }
  function addEditItem(p: Product) {
    setEditItems((prev) => [...prev, {
      product_id: p.id, product_name: getDeliveryName(p), sku: p.sku || undefined, category: p.category || undefined, quantity: 1, unit_price: getDeliveryPrice(p),
    }]);
  }
  function getDeliveryName(p: Product): string {
    const ce = catalogEntryMap.get(p.id);
    return ce?.delivery_name?.trim() || p.name;
  }

  // ALPS 송장 생성 (B2B 배송)
  const handleBookInvoice = async () => {
    if (!dl.customer_id) { toast.error('거래처 정보가 없어 송장을 생성할 수 없습니다'); return; }
    setBookingInvoice(true);
    try {
      // 고객 주소 조회
      const custRes = await fetch(`/api/customers/${dl.customer_id}`);
      const custData = await custRes.json();
      const cust = custData.customer;
      if (!cust?.address_road) { toast.error('거래처 주소가 등록되어 있지 않습니다'); return; }

      const gdsNm = items.map((it: DLItem) =>
        it.quantity > 1 ? `${it.product_name} x${it.quantity}` : it.product_name
      ).join(', ').slice(0, 750) || '마모루 제품';

      const res = await fetch('/api/lotte/book', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          deliveryId: dl.id,
          rcvName: dl.customer_name,
          rcvTel: cust.phone || '',
          rcvZip: cust.postcode || '',
          rcvAdr: `${cust.address_road || ''} ${cust.address_detail || ''}`.trim(),
          gdsNm,
        }),
      });
      const result = await res.json();
      if (!res.ok) { toast.error(result.error || '송장 생성 실패'); return; }
      toast.success(`송장 생성 완료: ${result.invNo}`);
      queryClient.invalidateQueries({ queryKey: ['delivery', deliveryId] });
      queryClient.invalidateQueries({ queryKey: ['deliveries'] });
    } catch (err) {
      toast.error('송장 생성 중 오류: ' + String(err));
    } finally {
      setBookingInvoice(false);
    }
  };
  const paymentStatus = (dl.payment_status as string) || 'unpaid';
  const dlVatType = ((dl.vat_type as string) || 'included') as 'included' | 'separate' | 'none';
  const totalAmount = (dl.total_amount as number) || 0;
  const discountAmount = (dl.discount_amount as number) || 0;
  const { supply, vat } = calcVAT(totalAmount, dlVatType);

  async function handleAction(action: string, extra?: Record<string, unknown>) {
    try {
      await updateDL.mutateAsync({ id: deliveryId, action, ...extra });
      queryClient.invalidateQueries({ queryKey: ['delivery', deliveryId] });
      setPendingAction(null);
    } catch {
      // error handled by hook
    }
  }

  function startEditing() {
    setEditing(true);
    setEditItems(items.map((i) => ({
      product_id: (i.product_id as string) || undefined,
      product_name: i.product_name as string,
      sku: (i.sku as string) || undefined,
      // 복원수리(category='RS') 등 카테고리 보존 — 누락 시 RS가 제품으로 둔갑해 집계 오류
      category: (i.category as string) || undefined,
      quantity: i.quantity as number,
      unit_price: i.unit_price as number,
    })));
    setEditMemo((dl.memo as string) || '');
    setEditExpectedDate((dl.expected_date as string) || '');
  }

  async function handleSaveEdit() {
    if (editItems.length === 0) { toast.error('품목을 추가해주세요'); return; }
    setSaving(true);
    try {
      await updateDL.mutateAsync({
        id: deliveryId,
        memo: editMemo,
        expected_date: editExpectedDate || undefined,
        // 품목 저장은 draft 에서만 (확정 후엔 재고/미수금 정합성 위해 품목 변경 막음 — 메모/날짜만 수정)
        ...(status === 'draft' ? { items: editItems } : {}),
      });
      queryClient.invalidateQueries({ queryKey: ['delivery', deliveryId] });
      queryClient.invalidateQueries({ queryKey: ['deliveries'] });
      setEditing(false);
      toast.success('납품서 수정 완료');
    } catch (err) {
      toast.error(String(err));
    } finally { setSaving(false); }
  }

  return (
    <div className="flex-1 overflow-y-auto space-y-4 p-4">
      {/* 헤더 */}
      <div className="flex items-center justify-between">
        <div>
          <h3 className="text-sm font-bold text-stone-900">{dl.dl_number as string}</h3>
          <p className="text-xs text-neutral-500 mt-0.5">
            {dl.customer_name as string} · {formatDate(dl.delivery_date as string)}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {status !== 'draft' && !dl.cancelled_at && (
            <Button variant="ghost" size="sm" onClick={() => setShowPrint(true)}>
              <Printer size={14} />납품서
            </Button>
          )}
          {status === 'draft' && !editing && (
            <Button variant="ghost" size="sm" onClick={startEditing}>
              <Pencil size={14} />편집
            </Button>
          )}
          {/* 110: 뱃지를 4단계로 (납품확정 → 출고대기 → 출고완료 → 배송완료). 규칙은 lib/deliveries/status.ts 단일출처 */}
          {(() => {
            const chip = getDeliveryStatusChip(dl);
            return <Badge className={chip.className}>{chip.label}</Badge>;
          })()}
        </div>
      </div>

      {/* 납품 정보 */}
      <Card>
        <div className="grid grid-cols-2 gap-3 text-sm">
          <div>
            <span className="text-xs text-neutral-500">거래처</span>
            <p className="font-semibold flex items-center gap-1.5">
              {dl.customer_name as string}
              {dl.customer_type ? (
                <Badge className={CUSTOMER_TYPE_COLOR[String(dl.customer_type)] || ''}>
                  {CUSTOMER_TYPE_LABEL[String(dl.customer_type)] || String(dl.customer_type)}
                </Badge>
              ) : null}
            </p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">납품일</span>
            <p>{formatDate(dl.delivery_date as string)}</p>
          </div>
          {dl.expected_date && (
            <div>
              <span className="text-xs text-neutral-500">납품 예정일</span>
              <p>{formatDate(dl.expected_date as string)}</p>
            </div>
          )}
          {/* 110: 출고일 — 집하 자동감지 건은 그 사실을 함께 표기 */}
          {dl.shipped_date ? (
            <div>
              <span className="text-xs text-neutral-500">출고일</span>
              <p>
                {formatDate(dl.shipped_date as string)}
                {dl.shipped_source === 'alps_pickup' && (
                  <span className="ml-1 text-[10px] text-neutral-400">· 기사님 수거 자동감지</span>
                )}
              </p>
            </div>
          ) : isAwaitingPickup(dl) ? (
            <div>
              <span className="text-xs text-neutral-500">출고</span>
              {/* 자세한 안내는 액션 카드가 맡는다 — 같은 문장을 두 번 쓰지 않는다 */}
              <p className="text-xs text-amber-600">출고대기</p>
            </div>
          ) : null}
          {dl.tracking_number && (
            <div>
              <span className="text-xs text-neutral-500">송장번호</span>
              <p className="font-mono text-xs">{dl.tracking_number as string}</p>
              {/* 150: 택배사 명시 — 롯데가 아니면 자동추적이 안 된다는 사실을 같이 보여준다 */}
              <p className="text-[11px] text-neutral-400">
                {courierLabel(dl.courier_name as string | null)}
                {!isAlpsTrackable(dl.courier_name as string | null) && ' · 자동 추적 미지원'}
              </p>
            </div>
          )}
          {/* 150: 택배 없이 직접 전달한 건 */}
          {!dl.tracking_number && isDirectHandover(dl.courier_name as string | null) && (
            <div>
              <span className="text-xs text-neutral-500">배송</span>
              <p className="text-xs">직접 전달 (송장 없음)</p>
            </div>
          )}
          {/* 배송완료 (ALPS 인수자등록 자동 감지) — 날짜+시간 */}
          {dl.delivered_at ? (
            <div>
              <span className="text-xs text-neutral-500">배송완료</span>
              <p className="text-green-600 font-medium">{formatDate(dl.delivered_at as string, 'M월 d일 HH:mm')}</p>
            </div>
          ) : null}
          <div>
            <span className="text-xs text-neutral-500">증빙유형</span>
            <p>{RECEIPT_LABEL[(dl.receipt_type as string)] || '미적용'}</p>
          </div>
        </div>
        {dl.memo && (
          <p className="mt-3 pt-3 border-t border-neutral-100 text-sm text-neutral-600">{dl.memo as string}</p>
        )}
      </Card>

      {/* 품목 */}
      {editing ? (
        <Card>
          <div className="flex items-center justify-between mb-2">
            <h4 className="text-xs font-semibold text-neutral-500">품목 편집</h4>
            <div className="flex gap-1">
              <Button variant="ghost" size="sm" onClick={() => setEditing(false)}><X size={14} />취소</Button>
              <Button size="sm" loading={saving} onClick={handleSaveEdit}><Save size={14} />저장</Button>
            </div>
          </div>

          {/* 메모 + 예정일 */}
          <div className="grid grid-cols-2 gap-2 mb-3">
            <div>
              <label className="text-xs text-neutral-500">납품 예정일</label>
              <input type="date" value={editExpectedDate} onChange={(e) => setEditExpectedDate(e.target.value)}
                className="w-full h-8 px-2 rounded border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">메모</label>
              <input type="text" value={editMemo} onChange={(e) => setEditMemo(e.target.value)} placeholder="메모"
                className="w-full h-8 px-2 rounded border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
          </div>

          {/* 편집 품목 */}
          <div className="space-y-2 mb-3">
            {editItems.map((item, idx) => (
              <div key={idx} className="flex flex-wrap items-center gap-2 py-1 border-b border-neutral-50 last:border-0">
                <div className="w-full sm:flex-1 sm:w-auto min-w-0">
                  <p className="text-xs font-medium truncate">{item.product_name}</p>
                </div>
                <div className="flex items-center gap-1">
                  <button onClick={() => setEditItems((prev) => prev.map((it, i) => i === idx ? { ...it, quantity: Math.max(1, it.quantity - 1) } : it))}
                    className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Minus size={10} /></button>
                  <span className="text-xs font-semibold w-6 text-center">{item.quantity}</span>
                  <button onClick={() => setEditItems((prev) => prev.map((it, i) => i === idx ? { ...it, quantity: it.quantity + 1 } : it))}
                    className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Plus size={10} /></button>
                </div>
                <input type="number" value={item.unit_price || ''} onChange={(e) => setEditItems((prev) => prev.map((it, i) => i === idx ? { ...it, unit_price: parseInt(e.target.value) || 0 } : it))}
                  className="w-20 h-7 px-2 rounded border border-neutral-200 bg-stone-50 text-xs text-right" />
                <button onClick={() => setEditItems((prev) => prev.filter((_, i) => i !== idx))}
                  className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500"><Trash2 size={10} /></button>
              </div>
            ))}
          </div>

          {/* 제품 추가 */}
          <div>
            <p className="text-xs text-neutral-400 mb-1">제품 추가</p>
            <input
              type="text"
              placeholder="제품명·SKU 검색..."
              value={addSearch}
              onChange={(e) => setAddSearch(e.target.value)}
              onKeyDown={addPick.onKeyDown}
              className="w-full h-8 px-3 mb-1.5 rounded-lg border border-neutral-200 text-xs placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-neutral-300"
            />
            <div className="max-h-[200px] overflow-y-auto border border-neutral-100 rounded-lg divide-y divide-neutral-50">
              {addableProducts.length === 0 && <p className="px-3 py-3 text-xs text-neutral-400 text-center">검색 결과 없음</p>}
              {addableProducts.map((p, pi) => (
                <button key={p.id}
                  onClick={() => addEditItem(p)}
                  onMouseEnter={() => addPick.setActiveIdx(pi)}
                  className={`w-full flex items-center justify-between px-3 py-2 text-xs transition text-left ${addPick.isActive(pi) ? 'bg-neutral-100' : 'hover:bg-neutral-50'}`}>
                  <span className="truncate font-medium">{getDeliveryName(p)}</span>
                  <span className="text-neutral-400 shrink-0 ml-2">{formatKRW(getDeliveryPrice(p))}</span>
                </button>
              ))}
            </div>
          </div>

          {/* 편집 합계 */}
          <div className="mt-3 pt-3 border-t border-neutral-200 flex justify-between text-sm font-bold">
            <span>합계</span>
            <span>{formatKRW(editItems.reduce((s, i) => s + i.quantity * i.unit_price, 0))}</span>
          </div>
        </Card>
      ) : (
        <Card>
          <h4 className="text-xs font-semibold text-neutral-500 mb-2">납품 품목</h4>
          <div className="space-y-2">
            {items.map((item) => (
              <div key={item.id as string} className="flex items-center justify-between py-1.5">
                <div>
                  <p className="text-sm font-medium">{item.product_name as string}</p>
                  <p className="text-xs text-neutral-500">
                    {item.sku && !String(item.sku).startsWith('IW-') && `${item.sku} · `}{formatKRW(item.unit_price as number)} x {item.quantity as number}
                  </p>
                </div>
                <span className="text-sm font-bold">{formatKRW(item.total_price as number)}</span>
              </div>
            ))}
          </div>
          <div className="mt-3 pt-3 border-t border-neutral-200 space-y-1">
            {discountAmount > 0 && (
              <div className="flex justify-between text-xs text-red-500">
                <span>할인</span><span>-{formatKRW(discountAmount)}</span>
              </div>
            )}
            <div className="flex justify-between text-sm font-bold">
              <span>합계</span>
              <span>{formatKRW(totalAmount)}</span>
            </div>
            <div className="flex justify-between text-xs text-neutral-500">
              <span>공급가액 {formatKRW(supply)}</span>
              {dlVatType !== 'none' && <span>부가세 {formatKRW(vat)}</span>}
              {dlVatType === 'none' && <span className="text-neutral-400">부가세 미적용</span>}
            </div>
          </div>
        </Card>
      )}

      {/* 결제 현황 */}
      <Card>
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">결제 현황</h4>
        <div className="grid grid-cols-3 gap-3 text-center">
          <div>
            <p className="text-xs text-neutral-400">총액</p>
            <p className="text-sm font-bold">{formatKRW(totalAmount)}</p>
          </div>
          <div>
            <p className="text-xs text-neutral-400">결제액</p>
            <p className="text-sm font-bold text-blue-600">{formatKRW((dl.paid_amount as number) || 0)}</p>
          </div>
          <div>
            <p className="text-xs text-neutral-400">미수금</p>
            {/* 납품 total_amount는 이미 net(할인 반영) → 할인 재차감 금지. 미수 = total - paid */}
            <p className={`text-sm font-bold ${totalAmount - ((dl.paid_amount as number) || 0) > 0 ? 'text-red-500' : 'text-green-600'}`}>
              {formatKRW(Math.max(0, totalAmount - ((dl.paid_amount as number) || 0)))}
            </p>
          </div>
        </div>
        <div className="flex justify-center mt-2">
          <Badge className={PAYMENT_COLOR[paymentStatus] || ''}>
            {PAYMENT_LABEL[paymentStatus] || paymentStatus}
          </Badge>
        </div>
      </Card>

      {/* 액션 — 표준 5슬롯 (components/ui/action-section.tsx)
           ① 지금 상태 한 줄 ② 주 액션(검은 버튼 1개) ③ 결제 ④ 드문 것·되돌리기는 접어둠 ⑤ 파괴적 액션
           2026-09-15 재설계: 전엔 가장 큰 검은 버튼이 제일 안 쓰는 [송장 재발급]이었고
           주 행동인 [출고 완료]는 흰 버튼으로 박스에 갇혀 있었으며,
           출고가 끝난 뒤에도 [송장 취소]만 빨갛게 덩그러니 남았다(사장님 지적).
           ⚠️ 슬롯 순서를 바꾸지 말 것 — 다른 상세 패널도 이 순서로 통일한다 */}
      {!dl.cancelled_at && status !== 'settled' && (
        <Card>
          <h4 className="text-xs font-semibold text-neutral-500 mb-2">액션</h4>
          <div className="space-y-2">

            {/* ① 지금 상태 */}
            {status === 'confirmed' && (
              dl.tracking_number ? (
                <ActionNote
                  tone={isAlpsTrackable(dl.courier_name as string | null) ? 'wait' : 'muted'}
                  title="출고대기"
                  meta={`${courierLabel(dl.courier_name as string | null)} ${dl.tracking_number as string}`}
                >
                  {isAlpsTrackable(dl.courier_name as string | null)
                    ? '롯데 기사님이 수거하면 자동으로 출고완료 처리됩니다 (1시간마다 확인)'
                    : '자동 감지가 안 되는 택배사입니다. 보내셨으면 아래 [출고 완료]를 눌러주세요.'}
                </ActionNote>
              ) : (
                <ActionNote tone="muted" title="납품확정" meta="송장 없음">
                  송장을 만들면 기사님 수거 시 자동으로 출고완료 처리됩니다.
                </ActionNote>
              )
            )}
            {status === 'shipped' && (
              <ActionNote
                tone="done"
                title="출고완료"
                meta={dl.tracking_number
                  ? `${courierLabel(dl.courier_name as string | null)} ${dl.tracking_number as string}`
                  : '직접 전달 (송장 없음)'}
              >
                {dl.delivered_at
                  ? '배송완료까지 확인됐습니다.'
                  : dl.tracking_number && isAlpsTrackable(dl.courier_name as string | null)
                    ? '인수자등록이 감지되면 배송완료로 바뀝니다 (1시간마다 확인)'
                    : '자동 배송추적은 되지 않습니다.'}
              </ActionNote>
            )}

            {/* ② 주 액션 — 딱 하나 */}
            {status === 'draft' && (
              <Button className="w-full" onClick={() => setPendingAction({
                action: 'confirm', label: '납품 확정',
                msg: '이 납품서를 확정합니다.\n확정 시 재고가 차감됩니다.',
              })} disabled={updateDL.isPending}>
                납품 확정 (재고 차감)
              </Button>
            )}
            {status === 'confirmed' && (
              dl.tracking_number ? (
                <Button className="w-full" onClick={() => setPendingAction({
                  action: 'ship', label: '출고 완료',
                  msg: `${courierLabel(dl.courier_name as string | null)} 송장 ${dl.tracking_number as string}(으)로 출고 처리합니다.`,
                  extra: { courier_name: (dl.courier_name as string | null) || COURIER_LOTTE },
                })} disabled={updateDL.isPending}>
                  출고 완료
                </Button>
              ) : (
                <Button className="w-full" onClick={handleBookInvoice} disabled={bookingInvoice || updateDL.isPending}>
                  {bookingInvoice ? '송장 생성 중...' : '🚚 송장 생성 (롯데택배)'}
                </Button>
              )
            )}

            {/* ③ 결제 — 상태와 무관하게 남은 할 일 */}
            {paymentStatus !== 'paid' && (
              <Button variant="secondary" className="w-full" onClick={() => setPendingAction({
                action: 'update_payment', label: '결제완료 처리',
                msg: `결제상태를 결제완료로 변경합니다.\n총액: ${formatKRW(totalAmount)}`,
                extra: { payment_status: 'paid' },
              })} disabled={updateDL.isPending}>
                결제완료 처리
              </Button>
            )}

            {/* ④ 드물게 쓰는 것 — 접어둔다 */}
            {status === 'confirmed' && (
              <MoreActions label={dl.tracking_number ? '다르게 보냈어요' : '다른 택배사로 보냈어요'}>
                  <div className="rounded-lg border border-neutral-200 p-2.5 space-y-1.5">
                    <div className="flex items-center gap-2">
                      <span className="w-12 shrink-0 text-[11px] text-neutral-500">택배사</span>
                      <select
                        value={courierInput}
                        onChange={(ev) => setCourierPick(ev.target.value)}
                        className="flex-1 h-9 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-neutral-300"
                      >
                        {COURIER_OPTIONS.map((c) => (
                          <option key={c} value={c}>{c === COURIER_DIRECT ? '직접 전달 (송장 없음)' : c}</option>
                        ))}
                      </select>
                    </div>
                    {courierInput !== COURIER_DIRECT && (
                      <div className="flex items-center gap-2">
                        <span className="w-12 shrink-0 text-[11px] text-neutral-500">송장</span>
                        <input
                          type="text"
                          value={trackingInput}
                          onChange={(ev) => setTrackingInput(ev.target.value)}
                          placeholder={dl.tracking_number ? '비우면 기존 송장 유지' : '송장번호'}
                          className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
                        />
                      </div>
                    )}
                    <Button variant="secondary" className="w-full" onClick={() => setPendingAction({
                      action: 'ship',
                      label: courierInput === COURIER_DIRECT ? '직접 전달 처리' : '출고 완료',
                      msg: courierInput === COURIER_DIRECT
                        ? '택배 없이 직접 전달한 것으로 출고 처리합니다.\n등록된 송장번호가 있으면 함께 지워집니다.'
                        : `${courierInput} · 송장 ${trackingInput || (dl.tracking_number as string) || '없음'}(으)로 출고 처리합니다.`,
                      extra: { tracking_number: trackingInput || undefined, courier_name: courierInput },
                    })} disabled={updateDL.isPending}>
                      {courierInput === COURIER_DIRECT ? '직접 전달로 출고 완료' : '이 방법으로 출고 완료'}
                    </Button>
                    {courierInput !== COURIER_LOTTE && courierInput !== COURIER_DIRECT && (
                      <p className="text-[10px] text-neutral-400 leading-relaxed">
                        롯데 외 택배사는 집하·배송완료 자동 추적이 되지 않습니다.
                      </p>
                    )}
                  </div>

                {/* 롯데 건에서만 재발급 의미가 있다 */}
                {dl.tracking_number && isAlpsTrackable(dl.courier_name as string | null) && (
                  <SubtleButton onClick={handleBookInvoice} disabled={bookingInvoice || updateDL.isPending}>
                    {bookingInvoice ? '송장 생성 중...' : '🚚 롯데 송장 재발급'}
                  </SubtleButton>
                )}

                {dl.tracking_number && (
                  <DangerLink
                    onClick={() => setPendingAction({
                      action: 'cancel_shipment', label: '송장 취소',
                      msg: `송장 ${dl.tracking_number as string}
(${courierLabel(dl.courier_name as string | null)})

먼저 롯데 ALPS 에서 집하취소를 완료하셨나요?
TMS 는 송장 기록만 지웁니다.
재고·매출·미수금은 그대로입니다.`,
                      variant: 'danger',
                    })}
                    disabled={updateDL.isPending}
                  >
                    송장 취소 (기록 지우기)
                  </DangerLink>
                )}
              </MoreActions>
            )}
            {status === 'shipped' && dl.tracking_number && (
              <MoreActions label="잘못 처리했나요?">
                <DangerLink
                  onClick={() => setPendingAction({
                    action: 'cancel_shipment', label: '송장 취소',
                    msg: `송장 ${dl.tracking_number as string}
(${courierLabel(dl.courier_name as string | null)})

송장 기록을 지우고 출고 전(납품확정)으로 되돌립니다.
롯데 송장이면 ALPS 집하취소를 먼저 해주세요.
재고·매출·미수금은 그대로입니다.`,
                    variant: 'danger',
                  })}
                  disabled={updateDL.isPending}
                >
                  송장 취소 · 출고 전으로 되돌리기
                </DangerLink>
              </MoreActions>
            )}

            {/* ⑤ 파괴적 액션 */}
            {(status === 'draft' || status === 'confirmed') && (
              <DangerZone>
                <DangerLink
                  onClick={() => setPendingAction({
                    action: 'cancel', label: '납품 취소',
                    msg: status === 'confirmed'
                      ? '이 납품을 취소합니다.\n확정된 납품은 재고가 복원됩니다.'
                      : '이 납품을 취소합니다.',
                    variant: 'danger',
                  })}
                  disabled={updateDL.isPending}
                >
                  납품 취소
                </DangerLink>
              </DangerZone>
            )}
          </div>
        </Card>
      )}

      {/* 취소 표시 */}
      {dl.cancelled_at && (
        <Card>
          <div className="text-center py-2">
            <Badge className="bg-red-100 text-red-600">취소됨</Badge>
            {dl.cancelled_reason && (
              <p className="text-xs text-neutral-500 mt-1">{dl.cancelled_reason as string}</p>
            )}
          </div>
        </Card>
      )}

      {/* 확인 모달 */}
      {pendingAction && (
        <ConfirmModal
          open={!!pendingAction}
          onClose={() => setPendingAction(null)}
          onConfirm={() => handleAction(pendingAction.action, pendingAction.extra)}
          title={pendingAction.label}
          message={<span className="whitespace-pre-wrap">{pendingAction.msg}</span>}
          confirmLabel={pendingAction.label}
          variant={pendingAction.variant || 'default'}
        />
      )}

      {/* 인쇄 모달 */}
      {showPrint && <DLPrintModal deliveryId={deliveryId} onClose={() => setShowPrint(false)} />}
    </div>
  );
}
