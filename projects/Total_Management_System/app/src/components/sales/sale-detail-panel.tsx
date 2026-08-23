'use client';

import { useState, useRef } from 'react';
import Link from 'next/link';
import { useQueryClient as __useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSale, useCancelSale, useReturnSale, useUpdatePaymentStatus, useUpdateSaleMemo, useEditSale, useRebuildSale, useProducts, useShipSale, useCancelSaleShipment, useMarkSaleShipped, useMarkSaleDelivered, useResendShipNotify, useMarkSalePacked } from '@/hooks/use-sales';
import { CustomerQuickModal } from '@/components/customers/customer-quick-modal';
import { CustomerNotes } from '@/components/shared/customer-notes';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { isB2BCustomerType } from '@/lib/sales/customer-type';
import { Hash, Ban, CheckCircle, AlertTriangle, Pencil, Save, FileText, Printer, Download, Truck, Package, ClipboardList, Copy, Link2 } from 'lucide-react';
import { PrepSheetModal } from './prep-sheet-modal';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { ReviewManagementCard } from '@/components/reviews/review-management-card';
import { LinkConsultationModal } from './link-consultation-modal';
import type { SaleChannel, OfflineSale, OfflineSaleItem, Product } from '@/lib/supabase/types';
import { getUnitPrice, getProductDisplayName, hasGroupPrice } from '@/lib/utils/pricing';
import { usePriceGroups } from '@/hooks/use-price-groups';
import { SerialPicker } from './serial-picker';
import { LabelPrintModal } from '@/components/labels/label-print-modal';
import { useLabelTemplate } from '@/hooks/use-label-templates';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합',
};

const PAYMENT_STATUS_COLOR: Record<string, string> = {
  paid: 'bg-green-100 text-green-700',
  unpaid: 'bg-red-100 text-red-700',
  partial: 'bg-yellow-100 text-yellow-700',
};

const PAYMENT_STATUS_LABEL: Record<string, string> = {
  paid: '결제완료', unpaid: '미결제', partial: '부분결제',
};

// 채널 칩 — 2026-07-17 4분류(매장/출장/톡/온라인) + 레거시 오프라인
const CHANNEL_CHIP: Record<string, { label: string; className: string }> = {
  store:   { label: '매장', className: 'bg-neutral-100 text-neutral-700' },
  field:   { label: '출장', className: 'bg-emerald-100 text-emerald-700' },
  talk:    { label: '톡',   className: 'bg-yellow-100 text-yellow-700' },
  online:  { label: '온라인(아임웹)', className: 'bg-blue-100 text-blue-700' },
  offline: { label: '오프라인', className: 'bg-neutral-100 text-neutral-500' }, // 레거시
};

/**
 * 운영 상태 칩 (2026-05-25 — 사장님 결정: '판매완료' 단일 통일 + '미수금' 우선)
 *   우선순위: 취소 > 반품 > 판매완료 > 배송중 > 배송대기 > (default 없음, 결제상태 칩이 별도 표시)
 *   미수금은 결제상태 칩에서 빨강으로 이미 강조됨 (중복 칩 X)
 *   2026-08-23: 주문관리와 통일 — 송장 발급·집하 전(invoice O·shipped_at X)에 '배송대기' 칩 추가
 */
function getSaleOperationChip(sale: OfflineSale): { label: string; className: string } | null {
  if (sale.cancelled_at) return { label: '취소', className: 'bg-red-100 text-red-700' };
  if (sale.returned_at) return { label: '반품', className: 'bg-violet-100 text-violet-700' };  // 종결 — 취소 다음 우선
  if (sale.delivered_at) return { label: '판매완료', className: 'bg-neutral-100 text-neutral-600' };
  if (sale.shipped_at)   return { label: '배송중', className: 'bg-green-100 text-green-700' };
  if (sale.invoice_number) return { label: '배송대기', className: 'bg-amber-100 text-amber-700' };  // 송장 발급, 집하 전
  return null;
}

interface Props {
  saleId: string;
}

export function SaleDetailPanel({ saleId }: Props) {
  const queryClient = __useQueryClient();
  const { data, isLoading } = useSale(saleId);
  const cancelSale = useCancelSale();
  const returnSale = useReturnSale();
  const updatePayment = useUpdatePaymentStatus();
  const updateMemo = useUpdateSaleMemo();
  const editSale = useEditSale();
  const rebuildSale = useRebuildSale();
  const shipSale = useShipSale();
  const cancelShipment = useCancelSaleShipment();
  const markShipped = useMarkSaleShipped();
  const markDelivered = useMarkSaleDelivered();
  const resendShipNotify = useResendShipNotify();
  const markPacked = useMarkSalePacked();   // 포장완료(준비완료) 토글 — 2026-07-18
  const [showPickupConfirm, setShowPickupConfirm] = useState(false);
  const [cancelMode, setCancelMode] = useState(false);
  const [cancelReason, setCancelReason] = useState('');
  const [returnMode, setReturnMode] = useState(false);
  const [returnReason, setReturnReason] = useState('');
  const [showPaidConfirm, setShowPaidConfirm] = useState(false);
  const [showCancelConfirm, setShowCancelConfirm] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const [showReceipt, setShowReceipt] = useState(false);
  const [showPrepSheet, setShowPrepSheet] = useState(false);
  const [showCustomer, setShowCustomer] = useState(false);
  const [editingMemo, setEditingMemo] = useState(false);
  const [memoValue, setMemoValue] = useState('');
  const [showShipConfirm, setShowShipConfirm] = useState(false);
  const [shipNotify, setShipNotify] = useState(true);
  const [showLinkConsultation, setShowLinkConsultation] = useState(false);
  const serialLabelTpl = useLabelTemplate('serial_40x20');
  const [labelSerial, setLabelSerial] = useState<{ product: string; serial: string } | null>(null);

  if (isLoading) {
    return (
      <div className="p-4 space-y-3">
        <Skeleton className="h-20 w-full" />
        <Skeleton className="h-32 w-full" />
      </div>
    );
  }

  if (!data) {
    return <p className="text-sm text-neutral-400 text-center py-8">판매 정보를 찾을 수 없습니다</p>;
  }

  const { sale, items, serials = [] } = data;
  const customerInfo = (data as { customerInfo?: { company_name: string | null; activity_name: string | null; position: string | null; postcode: string | null; address_road: string | null; address_detail: string | null } | null }).customerInfo ?? null;
  const hasStoreInfo = !!(customerInfo && (customerInfo.company_name || customerInfo.activity_name || customerInfo.position));
  const addressText = customerInfo
    ? [customerInfo.address_road, customerInfo.address_detail].filter(Boolean).join(' ').trim()
    : '';
  const s = sale as unknown as OfflineSale;
  const channel = CHANNEL_CHIP[(s.sale_channel || 'offline') as SaleChannel] || CHANNEL_CHIP.offline;

  const handleCancel = async () => {
    await cancelSale.mutateAsync({ id: saleId, reason: cancelReason });
    setCancelMode(false);
    setCancelReason('');
  };

  const handleMarkPaid = () => {
    updatePayment.mutate({
      id: saleId,
      payment_status: 'paid',
      // 실수납 = 소계 - 할인 (완납 시 할인 반영. total 그대로 넣으면 회계 과대계상)
      paid_amount: Math.max(0, s.total_amount - s.discount_amount),
    });
  };

  const handleSaveMemo = () => {
    updateMemo.mutate({ id: saleId, memo: memoValue });
    setEditingMemo(false);
  };

  return (
    <div className="p-4 space-y-4">
      {/* 헤더 — @container: 좌우배치를 '뷰포트' 아닌 '패널 폭' 기준으로 (그리드모드 420px 패널에서 리뷰카드 세로깨짐 방지) */}
      <div className="@container">
        <div className="flex items-center gap-2 mb-2 flex-wrap">
          <span className="text-sm font-bold text-stone-900">{s.sale_number}</span>
          <Badge className={PAYMENT_STATUS_COLOR[s.payment_status] || ''}>
            {PAYMENT_STATUS_LABEL[s.payment_status] || s.payment_status}
          </Badge>
          {/* 운영 상태 칩 (2026-05-25): 취소 / 판매완료 / 배송중 */}
          {(() => {
            const op = getSaleOperationChip(s);
            return op ? <Badge className={op.className}>{op.label}</Badge> : null;
          })()}
          <Badge className={channel.className}>{channel.label}</Badge>
        </div>
        {/* 2026-05-26: 헤더 정보 좌측 + 리뷰 관리 미니 우측 (사장님 시선 부담 ↓) */}
        {/* 좁은 패널(그리드모드 420px·모바일): 세로 스택 → 고객정보 폭 확보 / 넓은 패널: 좌우 (@md=28rem 기준) */}
        <div className="flex flex-col @md:flex-row @md:items-start gap-4">
          <div className="flex-1 grid grid-cols-2 gap-2 text-sm">
            <div>
              <span className="text-xs text-neutral-500">고객명</span>
              <p className="font-semibold">
                {s.customer_id ? (
                  <button onClick={() => setShowCustomer(true)} className="text-blue-600 hover:underline">{s.customer_name}</button>
                ) : s.customer_name}
              </p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">연락처</span>
              <p>{formatPhone(s.customer_phone) || '-'}</p>
            </div>
            {/* 매장명 · 활동명 · 직급 — 고객 정보 함께 표시 (값 있을 때만) */}
            {hasStoreInfo && (
              <div className="col-span-2 flex flex-wrap gap-x-5 gap-y-1 rounded-lg bg-neutral-50 px-3 py-2">
                <div>
                  <span className="text-xs text-neutral-500">매장명</span>
                  <p>{customerInfo!.company_name || '-'}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">활동명</span>
                  <p>{customerInfo!.activity_name || '-'}</p>
                </div>
                <div>
                  <span className="text-xs text-neutral-500">직급</span>
                  <p>{customerInfo!.position || '-'}</p>
                </div>
              </div>
            )}
            <div>
              <span className="text-xs text-neutral-500">판매일</span>
              <p>{formatDate(s.sale_date, 'yyyy.MM.dd')}</p>
            </div>
            {/* 주소지 (고객 등록 주소) — 2026-07-14 */}
            {addressText && (
              <div>
                <span className="text-xs text-neutral-500">주소지</span>
                <p className="leading-snug">
                  {addressText}
                  {customerInfo?.postcode && <span className="text-neutral-400 text-xs"> ({customerInfo.postcode})</span>}
                </p>
              </div>
            )}
            <div>
              <span className="text-xs text-neutral-500">결제방법</span>
              <p>{PAYMENT_METHOD_LABEL[s.payment_method] || s.payment_method}</p>
              {/* 복합 결제 상세 */}
              {s.payment_method === 'mixed' && (() => {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const pd = (s as any).payment_detail as Record<string, number> | null;
                if (!pd) return null;
                return (
                  <div className="mt-1 text-xs text-neutral-500 space-y-0.5">
                    {pd.card > 0 && <p>카드 {formatKRW(pd.card)}</p>}
                    {pd.cash > 0 && <p>현금 {formatKRW(pd.cash)}</p>}
                    {pd.transfer > 0 && <p>이체 {formatKRW(pd.transfer)}</p>}
                  </div>
                );
              })()}
            </div>
          </div>

          {/* 리뷰 관리 미니 — 취소건엔 표시 X */}
          {!s.cancelled_at && (
            <div className="w-full @md:w-auto @md:shrink-0 pt-1">
              <ReviewManagementCard
                source="sale"
                id={s.id}
                customerName={s.customer_name}
                customerPhone={s.customer_phone}
                promisedAt={(s as { review_promised_at?: string | null }).review_promised_at ?? null}
                promisedType={(s as { review_promised_type?: 'purchase' | 'repair' | 'consult' | null }).review_promised_type ?? null}
                promisedSubtype={(s as { review_promised_subtype?: string | null }).review_promised_subtype ?? null}
                requestSentAt={((s as { review_request_sent_at?: string | null }).review_request_sent_at) ?? ((s as { review_requested_at?: string | null }).review_requested_at ?? null)}
                submittedAt={(s as { review_submitted_at?: string | null }).review_submitted_at ?? null}
                hasRepairItem={data?.items.some((it) => String((it as Record<string, unknown>).category || '') === 'RS') ?? false}
                onChanged={() => queryClient.invalidateQueries({ queryKey: ['sale', saleId] })}
                compact={true}
                // 2026-05-26: 자동 발송 예정 판정용 sale 정보 전달
                shippedAt={(s as { shipped_at?: string | null }).shipped_at ?? null}
                deliveredAt={(s as { delivered_at?: string | null }).delivered_at ?? null}
                invoiceNumber={(s as { invoice_number?: string | null }).invoice_number ?? null}
              />
            </div>
          )}
        </div>
        {/* 메모 인라인 편집 */}
        <div className="mt-2 pt-2 border-t border-neutral-100">
          {editingMemo ? (
            <div className="flex gap-2 items-center">
              <input
                type="text"
                value={memoValue}
                onChange={(e) => setMemoValue(e.target.value)}
                placeholder="메모 입력"
                className="flex-1 h-8 px-2 rounded border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400"
                autoFocus
              />
              <button onClick={handleSaveMemo} className="p-1 hover:bg-neutral-100 rounded"><Save size={14} className="text-stone-900" /></button>
              <button onClick={() => setEditingMemo(false)} className="p-1 hover:bg-neutral-100 rounded text-xs text-neutral-500">취소</button>
            </div>
          ) : (
            <div className="flex items-center gap-1.5 text-sm text-neutral-600">
              <span>{s.memo || '메모 없음'}</span>
              <button onClick={() => { setMemoValue(s.memo || ''); setEditingMemo(true); }} className="p-0.5 hover:bg-neutral-100 rounded">
                <Pencil size={11} className="text-neutral-400" />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* 고객 상담 메모 (특징·불편·요구 등 — 고객 연결된 판매만) */}
      {s.customer_id && (
        <div className="border-t border-neutral-100 pt-3">
          <h3 className="text-xs font-bold text-stone-900 mb-2">고객 상담 메모</h3>
          <CustomerNotes customerId={s.customer_id} collapsed />
        </div>
      )}

      {/* 판매 항목 */}
      <div className="border-t border-neutral-100 pt-3">
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">판매 항목</h4>
        <div className="space-y-2">
          {(() => {
            // 시리얼-항목 매칭: sale_item_id 우선 → product_id → 순서 배분
            type Sr = { id: string; serial_number: string; product_id: string | null; sale_item_id?: string | null };
            const allSerials = serials as Sr[];
            const usedIds = new Set<string>();

            return (items as OfflineSaleItem[]).map((item) => {
              // 1순위: sale_item_id로 정확 매칭
              const bySaleItem = allSerials.filter((sr) => sr.sale_item_id === item.id && !usedIds.has(sr.id));
              bySaleItem.forEach((sr) => usedIds.add(sr.id));

              // 2순위: product_id 매칭 (sale_item_id 없는 경우)
              const byProduct = bySaleItem.length === 0
                ? allSerials.filter((sr) => sr.product_id && sr.product_id === item.product_id && !usedIds.has(sr.id))
                : [];
              byProduct.forEach((sr) => usedIds.add(sr.id));

              // 3순위: 나머지 unmatched 순서 배분
              const remaining = bySaleItem.length === 0 && byProduct.length === 0
                ? allSerials.filter((sr) => !sr.product_id && !sr.sale_item_id && !usedIds.has(sr.id)).slice(0, item.quantity)
                : [];
              remaining.forEach((sr) => usedIds.add(sr.id));

              const itemSerials = [...bySaleItem, ...byProduct, ...remaining];

              return (
                <div key={item.id} className="py-1.5">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-sm font-medium">{item.product_name}</p>
                      <p className="text-xs text-neutral-500">
                        {item.sku && !item.sku.startsWith('IW-') && `${item.sku} · `}{formatKRW(item.unit_price)} x {item.quantity}
                      </p>
                    </div>
                    <span className="text-sm font-bold">{formatKRW(item.total_price)}</span>
                  </div>
                  {itemSerials.length > 0 && (
                    <div className="mt-1 flex flex-wrap gap-1">
                      {itemSerials.map((sr) => (
                        <button key={sr.id} type="button"
                          onClick={() => setLabelSerial({ product: item.product_name, serial: sr.serial_number })}
                          title="시리얼 라벨 출력"
                          className="inline-flex items-center gap-0.5 text-[10px] font-mono bg-neutral-100 text-neutral-600 px-1.5 py-0.5 rounded hover:bg-stone-900 hover:text-white transition">
                          <Hash size={8} />{sr.serial_number}
                          <Printer size={9} className="ml-0.5 opacity-60" />
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              );
            });
          })()}
        </div>
      </div>

      {/* 합계 */}
      <div className="border-t border-neutral-200 pt-3 space-y-1">
        <div className="flex justify-between text-sm">
          <span className="text-neutral-500">소계</span>
          <span>{formatKRW(s.total_amount)}</span>
        </div>
        {s.discount_amount > 0 && (
          <div className="flex justify-between text-sm">
            <span className="text-neutral-500">할인</span>
            <span className="text-red-600">-{formatKRW(s.discount_amount)}</span>
          </div>
        )}
        <div className="flex justify-between text-sm font-bold">
          <span>결제 금액</span>
          <span className="text-stone-900">{formatKRW(s.paid_amount)}</span>
        </div>
        {s.payment_status !== 'paid' && (
          <div className="flex justify-between text-sm font-semibold text-red-500">
            <span>미수금</span>
            {/* 미수금 = 소계 - 할인 - 결제 (할인 반영. 리스트/통계/outstanding.ts와 동일 공식) */}
            <span>{formatKRW(Math.max(0, s.total_amount - s.discount_amount - s.paid_amount))}</span>
          </div>
        )}
        {s.supply_amount > 0 && (
          <div className="mt-2 pt-2 border-t border-dashed border-neutral-200 space-y-0.5">
            <div className="flex justify-between text-xs text-neutral-500">
              <span>공급가액</span>
              <span>{formatKRW(s.supply_amount)}</span>
            </div>
            <div className="flex justify-between text-xs text-neutral-500">
              <span>부가세 (10%)</span>
              <span>{formatKRW(s.vat_amount)}</span>
            </div>
          </div>
        )}
      </div>

      {/* 거래명세서 + 준비표 + 수정 */}
      <div className="flex gap-2">
        <button
          onClick={() => setShowReceipt(true)}
          className="flex-1 flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
        >
          <FileText size={14} />
          거래명세서
        </button>
        <button
          onClick={() => setShowPrepSheet(true)}
          className="flex-1 flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
        >
          <ClipboardList size={14} />
          준비표
        </button>
        {!s.cancelled_at && (
          <button
            onClick={() => setShowEditModal(true)}
            className="flex-1 flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
          >
            <Pencil size={14} />
            수정
          </button>
        )}
      </div>

      {/* 070: 출장/매장상담 link 정보 칩 — 사장님이 이 판매가 어떤 상담에서 시작됐는지 한눈에 확인 */}
      {!s.cancelled_at && data?.linkedConsultation && (
        <div className="flex items-center gap-1.5 text-xs text-neutral-500 px-1">
          <span>이 판매는 출장/매장상담</span>
          <Link
            href={`/consultations/${data.linkedConsultation.id}`}
            className="px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 font-mono font-semibold hover:bg-blue-100 transition"
          >
            {data.linkedConsultation.unique_id}
          </Link>
          <span>에서 시작됨</span>
        </div>
      )}

      {/* 2026-05-26: 리뷰 관리 카드 → 헤더 우측 compact 모드로 이동됨 (시각 부담 ↓) */}

      {/* 070: 수동 link 액션 — link 없는 sale + phone 있을 때만 노출 */}
      {!s.cancelled_at && !data?.linkedConsultation && s.customer_phone && (
        <button
          type="button"
          onClick={() => setShowLinkConsultation(true)}
          className="w-full flex items-center justify-center gap-1.5 text-[11px] text-neutral-500 hover:text-blue-600 py-1.5 transition"
        >
          <Link2 size={11} />
          이 판매를 출장/매장상담과 연결
        </button>
      )}

      {/* 070: link 있는 sale은 해제 옵션도 제공 (오링크 복구) */}
      {!s.cancelled_at && data?.linkedConsultation && (
        <button
          type="button"
          onClick={async () => {
            if (!window.confirm('상담 연결을 해제하시겠습니까? 후기 관리가 다시 이 판매에서 활성화됩니다.')) return;
            const res = await fetch(`/api/sales/${saleId}`, {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ action: 'unlink_consultation' }),
            });
            if (res.ok) {
              toast.success('연결 해제 완료');
              queryClient.invalidateQueries({ queryKey: ['sale', saleId] });
            } else {
              toast.error('해제 실패');
            }
          }}
          className="w-full flex items-center justify-center gap-1.5 text-[11px] text-neutral-400 hover:text-red-500 py-1.5 transition"
        >
          상담 연결 해제
        </button>
      )}

      {showLinkConsultation && (
        <LinkConsultationModal
          saleId={saleId}
          customerPhone={s.customer_phone}
          onClose={() => setShowLinkConsultation(false)}
          onLinked={() => {
            setShowLinkConsultation(false);
            queryClient.invalidateQueries({ queryKey: ['sale', saleId] });
          }}
        />
      )}

      {/* 액션 */}
      {s.cancelled_at ? (
        <div className="flex items-center gap-2 px-3 py-2.5 rounded-lg bg-red-50 border border-red-100">
          <Ban size={14} className="text-red-500 shrink-0" />
          <div className="text-xs">
            <p className="font-semibold text-red-700">취소됨 — {formatDate(s.cancelled_at)}</p>
            {s.cancelled_reason && <p className="text-red-600 mt-0.5">{s.cancelled_reason}</p>}
          </div>
        </div>
      ) : (s as Record<string, unknown>).returned_at ? (
        <div className="flex items-center gap-2 px-3 py-2.5 rounded-lg bg-purple-50 border border-purple-100">
          <Package size={14} className="text-purple-500 shrink-0" />
          <div className="text-xs">
            <p className="font-semibold text-purple-700">반품 완료 — {formatDate((s as Record<string, unknown>).returned_at as string)}</p>
            {(s as Record<string, unknown>).return_reason ? <p className="text-purple-600 mt-0.5">{String((s as Record<string, unknown>).return_reason)}</p> : null}
          </div>
        </div>
      ) : (
        <div className="flex items-center gap-2 pt-2 border-t border-neutral-100">
          {s.payment_status !== 'paid' && (
            <Button size="sm" onClick={() => setShowPaidConfirm(true)} disabled={updatePayment.isPending}>
              <CheckCircle size={14} />
              {updatePayment.isPending ? '처리 중...' : '결제완료로 변경'}
            </Button>
          )}
          <button onClick={() => setReturnMode(true)} className="text-xs text-purple-500 hover:text-purple-700 transition">
            반품 처리
          </button>
          <button onClick={() => setShowCancelConfirm(true)} className="text-xs text-red-500 hover:text-red-700 transition">
            판매 취소
          </button>
        </div>
      )}

      {/* 결제완료 확인 모달 */}
      <ConfirmModal
        open={showPaidConfirm}
        onClose={() => setShowPaidConfirm(false)}
        onConfirm={() => updatePayment.mutateAsync({ id: saleId, payment_status: 'paid', paid_amount: Math.max(0, s.total_amount - s.discount_amount) })}
        title="결제완료 처리"
        message={<>{s.customer_name}님의 결제를 완료 처리합니다.<br />금액: {formatKRW(Math.max(0, s.total_amount - s.discount_amount))}</>}
        confirmLabel="결제완료"
      />

      {/* 판매 취소 확인 모달 */}
      <ConfirmModal
        open={showCancelConfirm}
        onClose={() => { setShowCancelConfirm(false); setCancelReason(''); }}
        onConfirm={() => cancelSale.mutateAsync({ id: saleId, reason: cancelReason })}
        title="판매 취소"
        message={
          <div className="space-y-2">
            <div className="flex items-center gap-1.5 text-xs text-red-600">
              <AlertTriangle size={12} />
              <span>시리얼/재고가 복원됩니다. 이 작업은 되돌릴 수 없습니다.</span>
            </div>
            <input
              type="text"
              value={cancelReason}
              onChange={(e) => setCancelReason(e.target.value)}
              placeholder="취소 사유 (선택)"
              className="w-full h-8 px-3 rounded-lg border border-red-200 bg-red-50/50 text-sm focus:outline-none focus:ring-2 focus:ring-red-300"
            />
          </div>
        }
        confirmLabel="취소 확정"
        variant="danger"
      />

      {/* 반품 확인 모달 */}
      <ConfirmModal
        open={returnMode}
        onClose={() => { setReturnMode(false); setReturnReason(''); }}
        onConfirm={() => returnSale.mutateAsync({ id: saleId, reason: returnReason })}
        title="반품 처리"
        message={
          <div className="space-y-2">
            <div className="flex items-center gap-1.5 text-xs text-purple-600">
              <Package size={12} />
              <span>제품이 회수되고 재고/시리얼이 원래 창고로 복귀됩니다.</span>
            </div>
            <input
              type="text"
              value={returnReason}
              onChange={(e) => setReturnReason(e.target.value)}
              placeholder="반품 사유 (예: 단순변심, 불량 등)"
              className="w-full h-8 px-3 rounded-lg border border-purple-200 bg-purple-50/50 text-sm focus:outline-none focus:ring-2 focus:ring-purple-300"
            />
          </div>
        }
        confirmLabel="반품 확정"
        variant="danger"
      />

      {/* 택배 발송 / 매장 수령 */}
      {!s.cancelled_at && (
        <div className="pt-2 border-t border-neutral-100">
          {/* 포장(준비) 완료 — 2026-07-18. 송장 유무와 무관하게 표시. 출고/수령되면 숨김(이미 끝난 단계) */}
          {!s.shipped_at && !s.delivered_at && (
            <div className="mb-3">
              {s.packed_at ? (
                <div className="flex items-center justify-between gap-2 px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-200">
                  <span className="text-xs font-semibold text-emerald-700 flex items-center gap-1.5">
                    <CheckCircle size={13} />
                    준비완료 · 포장 끝 {formatDate(s.packed_at, 'M월 d일 HH:mm')}
                  </span>
                  <button
                    onClick={() => markPacked.mutate({ ids: [saleId], packed: false })}
                    disabled={markPacked.isPending}
                    className="text-[11px] text-neutral-400 hover:text-neutral-600 underline shrink-0 disabled:opacity-50"
                  >
                    해제
                  </button>
                </div>
              ) : (
                <button
                  onClick={() => markPacked.mutate({ ids: [saleId], packed: true })}
                  disabled={markPacked.isPending}
                  className="w-full flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition disabled:opacity-50"
                >
                  <Package size={14} />
                  {markPacked.isPending ? '처리 중...' : '포장완료 (준비완료로 표시)'}
                </button>
              )}
            </div>
          )}
          {s.invoice_number ? (
            <div className="space-y-2">
              <div className="flex items-center gap-2">
                <Package size={14} className="text-green-600" />
                <span className="text-sm font-mono font-medium">{s.invoice_number}</span>
                <span className="text-xs text-neutral-400">{s.courier_name || '롯데택배'}</span>
              </div>
              {s.shipped_at ? (
                <>
                  <p className="text-xs text-green-600 flex items-center gap-1 flex-wrap">
                    <CheckCircle size={12} />
                    출고완료 {formatDate(s.shipped_at)}
                    {/* 109: 롯데 기사님 수거 스캔으로 자동 처리된 건 */}
                    {s.shipped_source === 'alps_pickup' && (
                      <span className="text-[10px] text-neutral-400">· 기사님 수거 자동감지</span>
                    )}
                  </p>
                  {/* 109: 출고 알림톡 발송 여부 + 110/버그수정: 미발송 시 수동 발송 버튼 */}
                  {s.shipped_notified_at ? (
                    <p className="text-[11px] text-neutral-400">
                      출고 알림톡 발송됨 {formatDate(s.shipped_notified_at, 'M월 d일 HH:mm')}
                    </p>
                  ) : isB2BCustomerType(s.customer_type) ? (
                    <p className="text-[11px] text-neutral-400">거래처(B2B) — 출고 알림톡 미발송</p>
                  ) : s.customer_phone ? (
                    <div className="flex items-center gap-2 pt-0.5">
                      <span className="text-[11px] text-amber-600">출고 알림톡 미발송</span>
                      <button
                        onClick={() => resendShipNotify.mutate({ id: saleId })}
                        disabled={resendShipNotify.isPending}
                        className="text-[11px] font-medium text-neutral-700 underline hover:text-neutral-900 disabled:opacity-50"
                      >
                        {resendShipNotify.isPending ? '발송 중…' : '출고 알림톡 보내기'}
                      </button>
                    </div>
                  ) : null}
                  {/* 배송완료 표시 (자동 cron OR 수동 fallback) */}
                  {s.delivered_at ? (
                    <p className="text-xs text-neutral-500 flex items-center gap-1">
                      <CheckCircle size={12} className="text-neutral-400" />
                      배송완료 {formatDate(s.delivered_at, 'M월 d일 HH:mm')}
                    </p>
                  ) : (
                    <div className="space-y-1 pt-1">
                      <p className="text-xs text-neutral-400 leading-relaxed">
                        ALPS 인수자등록 자동 감지 시 배송완료 전환됩니다 (1시간마다 자동 확인)
                      </p>
                      <button
                        onClick={() => markDelivered.mutate({ id: saleId, mode: 'delivery' })}
                        disabled={markDelivered.isPending}
                        className="text-xs text-neutral-500 hover:text-neutral-700 underline transition disabled:opacity-50"
                      >
                        수동 배송완료 처리
                      </button>
                    </div>
                  )}
                </>
              ) : (
                <>
                  {/* 109: 자동 감지 안내 — "왜 버튼 안 눌렀는데 됐지?" 를 막는 문구 */}
                  <p className="text-[11px] text-neutral-400 leading-relaxed">
                    롯데 기사님이 수거하면 자동으로 출고완료 처리됩니다 (1시간마다 확인)
                    {!isB2BCustomerType(s.customer_type) && ' · 출고 알림톡도 자동 발송'}
                  </p>
                  <Button
                    size="sm"
                    onClick={() => { setShipNotify(true); setShowShipConfirm(true); }}
                    disabled={markShipped.isPending}
                    className="w-full"
                  >
                    <Truck size={14} />
                    {markShipped.isPending ? '처리 중...' : '출고완료'}
                  </Button>
                  <button onClick={() => cancelShipment.mutate(saleId)}
                    disabled={cancelShipment.isPending}
                    className="w-full text-center text-xs text-red-400 hover:text-red-600">
                    {cancelShipment.isPending ? '취소 중...' : '송장 취소'}
                  </button>
                </>
              )}
            </div>
          ) : (
            // 송장 없음 — 택배 발송 OR 매장 수령 두 옵션
            <div className="space-y-2">
              {s.delivered_at ? (
                <p className="text-xs text-neutral-500 flex items-center gap-1">
                  <CheckCircle size={12} className="text-neutral-400" />
                  고객 수령 완료 {formatDate(s.delivered_at)}
                </p>
              ) : (
                <>
                  <button
                    onClick={() => shipSale.mutate(saleId)}
                    disabled={shipSale.isPending}
                    className="w-full flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
                  >
                    <Truck size={14} />
                    {shipSale.isPending ? '송장 생성 중...' : '택배 발송 (송장 생성)'}
                  </button>
                  <button
                    onClick={() => setShowPickupConfirm(true)}
                    disabled={markDelivered.isPending}
                    className="w-full flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition disabled:opacity-50"
                  >
                    <Package size={14} />
                    {markDelivered.isPending ? '처리 중...' : '고객 수령 완료'}
                  </button>
                </>
              )}
            </div>
          )}
        </div>
      )}

      {/* 고객 수령 완료 확인 모달 */}
      <ConfirmModal
        open={showPickupConfirm}
        onClose={() => setShowPickupConfirm(false)}
        onConfirm={() => {
          markDelivered.mutate({ id: saleId, mode: 'pickup' });
          setShowPickupConfirm(false);
        }}
        title="고객 수령 완료"
        message={
          <div className="space-y-2">
            <p>{s.customer_name} 고객이 매장에서 제품을 직접 받아가신 것으로 처리합니다.</p>
            <p className="text-xs text-neutral-500">송장 없이 판매완료 상태로 전환됩니다.</p>
          </div>
        }
        confirmLabel="수령 완료"
      />


      {/* 출고완료 확인 모달 */}
      <ConfirmModal
        open={showShipConfirm}
        onClose={() => setShowShipConfirm(false)}
        onConfirm={() => {
          markShipped.mutate({ id: saleId, send_notification: shipNotify && !!s.customer_phone });
          setShowShipConfirm(false);
        }}
        title="출고 완료"
        message={
          <div className="space-y-3">
            <p>
              송장 <strong className="font-mono">{(s as Record<string, unknown>).invoice_number as string}</strong>으로
              출고 완료 처리합니다.
            </p>
            <label className={`flex items-center gap-2 ${s.customer_phone ? 'cursor-pointer' : 'opacity-50 cursor-not-allowed'}`}>
              <input
                type="checkbox"
                checked={shipNotify && !!s.customer_phone}
                onChange={(e) => setShipNotify(e.target.checked)}
                disabled={!s.customer_phone}
                className="w-4 h-4 rounded border-neutral-300"
              />
              <span className="text-sm text-neutral-600">알림톡 함께 보내기</span>
            </label>
            {!s.customer_phone && (
              <p className="text-xs text-amber-600">
                고객 연락처가 없어 알림톡을 보낼 수 없습니다
              </p>
            )}
          </div>
        }
        confirmLabel={shipNotify && s.customer_phone ? '출고완료 + 알림톡' : '출고완료'}
      />

      {/* 고객 퀵뷰 모달 */}
      {s.customer_id && (
        <CustomerQuickModal customerId={s.customer_id} open={showCustomer} onClose={() => setShowCustomer(false)} />
      )}

      {/* 거래명세서 모달 */}
      {showReceipt && data && (
        <ReceiptModal
          sale={s}
          items={data.items}
          customerType={(s as Record<string, unknown>).customer_type as string | undefined}
          onClose={() => setShowReceipt(false)}
        />
      )}

      {showPrepSheet && data && (
        <PrepSheetModal
          saleIds={[saleId]}
          preloaded={{ sourceType: 'sale', sale: s, items: data.items, serials: data.serials || [] }}
          onClose={() => setShowPrepSheet(false)}
        />
      )}

      {/* 판매 수정 모달 */}
      {showEditModal && data && (
        <FullEditSaleModal
          sale={s}
          items={data.items}
          serials={serials}
          saleId={saleId}
          onClose={() => setShowEditModal(false)}
          rebuildSale={rebuildSale}
        />
      )}

      {/* 시리얼 라벨 출력 (판매 직후 — 각 시리얼 칩 클릭) */}
      {labelSerial && (
        <LabelPrintModal
          template={serialLabelTpl}
          data={{ product: labelSerial.product, serial: labelSerial.serial }}
          title="시리얼 라벨"
          onClose={() => setLabelSerial(null)}
        />
      )}
    </div>
  );
}

/** 판매 전체 수정 모달 (제품 추가/삭제 + 금액/결제 + 시리얼 수정) */
function FullEditSaleModal({ sale, items: originalItems, serials: existingSerials, saleId, onClose, rebuildSale }: {
  sale: OfflineSale;
  items: OfflineSaleItem[];
  serials: Array<{ id: string; serial_number: string; product_id: string | null; sale_item_id?: string | null }>;
  saleId: string;
  onClose: () => void;
  rebuildSale: ReturnType<typeof useRebuildSale>;
}) {
  const { data: products = [] } = useProducts();
  const priceGroups = usePriceGroups();

  // 기존 시리얼을 아이템별로 매칭하여 프리필 (객체 + ID 모두 보관 — 2026-05-18 fix)
  // SerialPicker UI에 "현재 등록된 시리얼" 표시 + 해제 버튼을 위해 객체 필요
  const buildInitialSerials = () => {
    const used = new Set<string>();
    return originalItems.map((item) => {
      // 1순위: sale_item_id 매칭
      const bySaleItem = existingSerials.filter((sr) => sr.sale_item_id === item.id && !used.has(sr.id));
      bySaleItem.forEach((sr) => used.add(sr.id));
      // 2순위: product_id 매칭
      const byProduct = bySaleItem.length === 0
        ? existingSerials.filter((sr) => sr.product_id && sr.product_id === item.product_id && !used.has(sr.id))
        : [];
      byProduct.forEach((sr) => used.add(sr.id));
      return [...bySaleItem, ...byProduct].map((sr) => ({ id: sr.id, serial_number: sr.serial_number }));
    });
  };
  const initialSerials = buildInitialSerials();
  const initialSerialIds = initialSerials.map((arr) => arr.map((sr) => sr.id));

  const [editItems, setEditItems] = useState(
    originalItems.map((it, idx) => ({
      product_id: it.product_id || undefined,
      product_name: it.product_name,
      sku: it.sku || undefined,
      category: ((it as Record<string, unknown>).category as string) || undefined,
      quantity: it.quantity,
      unit_price: it.unit_price,
      total_price: it.total_price,
      serial_ids: initialSerialIds[idx] || [] as string[],
      manualSerials: [] as string[],
    }))
  );
  // Phase A — 시리얼 다른 판매에서 이전 동의 플래그 (2026-05-18)
  const [allowSerialTransfer, setAllowSerialTransfer] = useState(false);
  const [paymentMethod, setPaymentMethod] = useState<string>(sale.payment_method);
  // 복합 결제 분리 금액 — 기존 payment_detail 로 초기화(재수정 시 값 유지)
  const [mixedCard, setMixedCard] = useState(sale.payment_detail?.card || 0);
  const [mixedCash, setMixedCash] = useState(sale.payment_detail?.cash || 0);
  const [mixedTransfer, setMixedTransfer] = useState(sale.payment_detail?.transfer || 0);
  const [paymentStatus, setPaymentStatus] = useState<string>(sale.payment_status || 'paid');
  const [paidAmount, setPaidAmount] = useState(sale.paid_amount || 0);
  const [saleChannel, setSaleChannel] = useState<string>((sale as Record<string, unknown>).sale_channel as string || 'offline');
  const [discountAmount, setDiscountAmount] = useState(sale.discount_amount || 0);
  const [saleDate, setSaleDate] = useState(sale.sale_date);
  const [editMemo, setEditMemo] = useState(sale.memo || '');
  const [productSearch, setProductSearch] = useState('');
  const [customName, setCustomName] = useState('');
  const [customPrice, setCustomPrice] = useState('');

  const totalAmount = editItems.reduce((s, it) => s + it.unit_price * it.quantity, 0);
  const finalAmount = totalAmount - discountAmount;

  const removeItem = (idx: number) => setEditItems((prev) => prev.filter((_, i) => i !== idx));
  const updateQty = (idx: number, delta: number) => {
    setEditItems((prev) => prev.map((it, i) => i === idx ? { ...it, quantity: Math.max(1, it.quantity + delta), total_price: it.unit_price * Math.max(1, it.quantity + delta) } : it));
  };

  const customerType = (sale as Record<string, unknown>).customer_type as string | undefined;

  const addProduct = (p: Product) => {
    const unitPrice = getUnitPrice(p, customerType, priceGroups);
    const displayName = getProductDisplayName(p, customerType, priceGroups);
    const existing = editItems.findIndex((it) => it.product_id === p.id);
    if (existing >= 0) {
      updateQty(existing, 1);
    } else {
      setEditItems((prev) => [...prev, {
        product_id: p.id,
        product_name: displayName,
        sku: p.sku,
        category: p.category || undefined,
        quantity: 1,
        unit_price: unitPrice,
        total_price: unitPrice,
        serial_ids: [],
        manualSerials: [],
      }]);
    }
    setProductSearch('');
  };

  const addCustomProduct = () => {
    const name = customName.trim();
    const price = parseInt(customPrice) || 0;
    if (!name || price <= 0) return;
    setEditItems((prev) => [...prev, {
      product_id: undefined,
      product_name: name,
      sku: undefined,
      category: undefined,
      quantity: 1,
      unit_price: price,
      total_price: price,
      serial_ids: [],
      manualSerials: [],
    }]);
    setCustomName('');
    setCustomPrice('');
  };

  const filteredProducts = productSearch.length >= 1
    ? products.filter((p) => p.name.toLowerCase().includes(productSearch.toLowerCase()) || (p.sku || '').toLowerCase().includes(productSearch.toLowerCase())).slice(0, 8)
    : [];

  const handleSave = async () => {
    // 복합결제 = 받은 돈을 수단별로 쪼갠 것 → 합계가 곧 수납액. 결제상태/미수 자동 판정(회계 정합)
    const mixedTotal = mixedCard + mixedCash + mixedTransfer;
    const isMixedPay = paymentMethod === 'mixed';
    const effStatus = isMixedPay ? (mixedTotal >= finalAmount ? 'paid' : mixedTotal <= 0 ? 'unpaid' : 'partial') : paymentStatus;
    const effPaid = isMixedPay ? Math.min(mixedTotal, finalAmount) : (paymentStatus === 'paid' ? finalAmount : paymentStatus === 'unpaid' ? 0 : paidAmount);
    await rebuildSale.mutateAsync({
      id: saleId,
      items: editItems.map((it) => ({
        ...it,
        total_price: it.unit_price * it.quantity,
        manual_serials: it.manualSerials?.length ? it.manualSerials : undefined,
      })),
      sale_info: {
        total_amount: totalAmount,
        discount_amount: discountAmount,
        payment_method: paymentMethod,
        payment_detail: paymentMethod === 'mixed' ? { card: mixedCard, cash: mixedCash, transfer: mixedTransfer } : undefined,
        payment_status: effStatus,
        paid_amount: effPaid,
        sale_date: saleDate,
        memo: editMemo.trim() || undefined,
        sale_channel: saleChannel,
      },
      allow_serial_transfer: allowSerialTransfer, // Phase A — 사장님 명시 동의 시 강탈 허용
    });
    onClose();
  };

  const METHODS = [
    { value: 'card', label: '카드' },
    { value: 'cash', label: '현금' },
    { value: 'transfer', label: '이체' },
    { value: 'mixed', label: '복합' },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-[95vw] max-w-[550px] max-h-[90vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-3 border-b border-neutral-200 flex items-center justify-between">
          <h3 className="text-sm font-bold text-neutral-800">판매 수정</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg">×</button>
        </div>

        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
          {/* 항목 목록 */}
          <div>
            <label className="text-xs font-semibold text-neutral-600 mb-2 block">품목 ({editItems.length})</label>
            <div className="space-y-2">
              {editItems.map((it, idx) => (
                <div key={idx} className="p-2 rounded-lg border border-neutral-200 bg-neutral-50">
                  <div className="flex items-center gap-2">
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium truncate">{it.product_name}</p>
                      <p className="text-xs text-neutral-500">{formatKRW(it.unit_price)} × {it.quantity} = {formatKRW(it.unit_price * it.quantity)}</p>
                    </div>
                    <div className="flex items-center gap-1 shrink-0">
                      <button onClick={() => updateQty(idx, -1)} className="w-6 h-6 rounded bg-neutral-200 text-xs font-bold">−</button>
                      <input
                        type="number"
                        min="1"
                        value={it.quantity}
                        onChange={(e) => {
                          const val = parseInt(e.target.value) || 1;
                          setEditItems((prev) => prev.map((item, i) => i === idx ? { ...item, quantity: Math.max(1, val), total_price: item.unit_price * Math.max(1, val) } : item));
                        }}
                        className="w-10 h-6 text-center text-xs font-bold border border-neutral-200 rounded bg-white focus:outline-none [appearance:textfield] [&::-webkit-inner-spin-button]:appearance-none"
                      />
                      <button onClick={() => updateQty(idx, 1)} className="w-6 h-6 rounded bg-neutral-200 text-xs font-bold">+</button>
                      <button onClick={() => removeItem(idx)} className="w-6 h-6 rounded bg-red-100 text-red-500 text-xs font-bold ml-1">×</button>
                    </div>
                  </div>
                  {/* 시리얼 피커 — 모든 아이템 (임시제품 포함)
                       currentSerials: 이 판매에 이미 등록된 시리얼 (sold 상태) — 해제 가능 표시 (2026-05-18 fix)
                       currentSaleId: 중복 검증 시 같은 판매 시리얼은 충돌로 보지 않음 (Phase A 2026-05-18) */}
                  <SerialPicker
                    productId={it.product_id || ''}
                    quantity={it.quantity}
                    selectedSerialIds={it.serial_ids}
                    onSelect={(ids) => setEditItems((prev) => prev.map((item, i) => i === idx ? { ...item, serial_ids: ids } : item))}
                    manualSerials={it.manualSerials}
                    onManualSerialsChange={(serials) => setEditItems((prev) => prev.map((item, i) => i === idx ? { ...item, manualSerials: serials } : item))}
                    currentSerials={initialSerials[idx] || []}
                    currentSaleId={saleId}
                    onTransferConsent={() => setAllowSerialTransfer(true)}
                    reservedSerials={editItems.flatMap((e, i) => i === idx ? [] : (e.manualSerials || []))}
                  />
                </div>
              ))}
            </div>
          </div>

          {/* 제품 추가 검색 */}
          <div>
            <label className="text-xs font-semibold text-neutral-600 mb-1 block">제품 추가</label>
            <input
              type="text"
              value={productSearch}
              onChange={(e) => setProductSearch(e.target.value)}
              placeholder="제품명 또는 SKU 검색"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400"
            />
            {filteredProducts.length > 0 && (
              <div className="mt-1 border border-neutral-200 rounded-lg max-h-40 overflow-y-auto divide-y divide-neutral-100">
                {filteredProducts.map((p) => {
                  const displayPrice = getUnitPrice(p, customerType, priceGroups);
                  const displayName = getProductDisplayName(p, customerType, priceGroups);
                  const isB2B = hasGroupPrice(p, customerType, priceGroups);
                  return (
                    <button key={p.id} onClick={() => addProduct(p)}
                      className="w-full text-left px-3 py-2 hover:bg-neutral-50 transition">
                      <span className="text-sm font-medium">{displayName}</span>
                      <span className="text-xs text-neutral-500 ml-2">{p.sku}</span>
                      <span className="float-right">
                        {isB2B && (
                          <span className="text-[11px] text-neutral-400 line-through mr-1">{formatKRW(p.price)}</span>
                        )}
                        <span className="text-xs font-bold text-neutral-700">{formatKRW(displayPrice)}</span>
                      </span>
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          {/* 임시 제품 추가 */}
          <div>
            <label className="text-xs font-semibold text-neutral-600 mb-1 block">임시 제품</label>
            <div className="flex gap-2">
              <input
                type="text"
                value={customName}
                onChange={(e) => setCustomName(e.target.value)}
                placeholder="제품명"
                className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400"
              />
              <input
                type="number"
                value={customPrice}
                onChange={(e) => setCustomPrice(e.target.value)}
                placeholder="단가"
                className="w-24 h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400"
                onKeyDown={(e) => { if (e.key === 'Enter') addCustomProduct(); }}
              />
              <button
                onClick={addCustomProduct}
                disabled={!customName.trim() || !customPrice || parseInt(customPrice) <= 0}
                className="px-3 h-9 rounded-lg bg-neutral-900 text-white text-xs font-medium disabled:opacity-30"
              >추가</button>
            </div>
          </div>

          {/* 판매일 + 결제 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">판매일</label>
              <input type="date" value={saleDate} max={new Date().toISOString().slice(0, 10)}
                onChange={(e) => setSaleDate(e.target.value)}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">할인</label>
              <input type="number" value={discountAmount}
                onChange={(e) => setDiscountAmount(Number(e.target.value))}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
          </div>

          <div>
            <label className="text-xs text-neutral-500 mb-1 block">결제방법</label>
            <div className="flex gap-2">
              {METHODS.map((m) => (
                <button key={m.value} onClick={() => setPaymentMethod(m.value)}
                  className={`flex-1 py-1.5 text-xs rounded-md border transition ${paymentMethod === m.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                >{m.label}</button>
              ))}
            </div>
            {/* 복합 결제 분리 입력 — 등록 폼과 동일(카드/현금/이체 + 합계검증) */}
            {paymentMethod === 'mixed' && (
              <div className="mt-2 p-2 rounded-lg bg-neutral-50 border border-neutral-200 space-y-1.5">
                {(['카드', '현금', '이체'] as const).map((label, i) => {
                  const val = [mixedCard, mixedCash, mixedTransfer][i];
                  const setter = [setMixedCard, setMixedCash, setMixedTransfer][i];
                  return (
                    <div key={label} className="flex items-center gap-1.5">
                      <span className="text-[10px] text-neutral-500 w-8">{label}</span>
                      <input type="number" min={0} value={val || ''} onChange={(e) => setter(parseInt(e.target.value) || 0)}
                        className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs text-right bg-white focus:outline-none focus:ring-1 focus:ring-stone-400" placeholder="0" />
                    </div>
                  );
                })}
                {(() => { const t = mixedCard + mixedCash + mixedTransfer; const short = finalAmount - t; return (
                  <div className={`flex justify-between text-[10px] font-semibold pt-1 border-t border-neutral-200 ${short === 0 ? 'text-green-600' : short > 0 ? 'text-red-500' : 'text-orange-500'}`}>
                    <span>합계 (= 받은 금액)</span><span>{formatKRW(t)}{short === 0 ? ' ✓ 완납' : short > 0 ? ` · 미수 ${formatKRW(short)}` : ` · 초과 ${formatKRW(-short)}`}</span>
                  </div>
                ); })()}
              </div>
            )}
          </div>

          {/* 결제상태 — 복합결제는 합계가 자동 결정하므로 숨김 */}
          {paymentMethod !== 'mixed' && (
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">결제상태</label>
            <div className="flex gap-2">
              {[{ value: 'paid', label: '결제완료' }, { value: 'partial', label: '부분결제' }, { value: 'unpaid', label: '미결제' }].map((s) => (
                <button key={s.value} onClick={() => setPaymentStatus(s.value)}
                  className={`flex-1 py-1.5 text-xs rounded-md border transition ${paymentStatus === s.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                >{s.label}</button>
              ))}
            </div>
            {paymentStatus === 'partial' && (
              <input type="number" value={paidAmount} onChange={(e) => setPaidAmount(Number(e.target.value))}
                placeholder="입금액" className="w-full h-9 px-3 mt-2 rounded-lg border border-neutral-200 text-sm" />
            )}
          </div>
          )}

          {/* 판매채널 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">판매채널</label>
            <div className="flex gap-2">
              {[{ value: 'store', label: '매장' }, { value: 'field', label: '출장' }, { value: 'talk', label: '톡' }, { value: 'online', label: '온라인' }].map((c) => (
                <button key={c.value} onClick={() => setSaleChannel(c.value)}
                  className={`flex-1 py-1.5 text-xs rounded-md border transition ${saleChannel === c.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                >{c.label}</button>
              ))}
            </div>
          </div>

          {/* 메모 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">메모</label>
            <input type="text" value={editMemo} onChange={(e) => setEditMemo(e.target.value)}
              placeholder="메모" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
          </div>

          {/* 합계 */}
          <div className="bg-neutral-50 rounded-lg p-3 text-sm">
            <div className="flex justify-between"><span>합계</span><span className="font-bold">{formatKRW(totalAmount)}</span></div>
            {discountAmount > 0 && <div className="flex justify-between text-neutral-500"><span>할인</span><span>-{formatKRW(discountAmount)}</span></div>}
            <div className="flex justify-between font-bold text-base pt-1 border-t border-neutral-200 mt-1">
              <span>결제금액</span><span>{formatKRW(finalAmount)}</span>
            </div>
          </div>

          <p className="text-[11px] text-red-500">⚠ 저장 시 기존 시리얼/재고가 복원된 후 새 구성으로 재적용됩니다.</p>
        </div>

        <div className="px-5 py-3 border-t border-neutral-200 flex gap-2">
          <button onClick={onClose} className="flex-1 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600">취소</button>
          <button
            onClick={handleSave}
            disabled={rebuildSale.isPending || editItems.length === 0}
            className="flex-1 py-2 rounded-lg bg-neutral-900 text-white text-sm font-medium disabled:opacity-50"
          >
            {rebuildSale.isPending ? '저장 중...' : '저장'}
          </button>
        </div>
      </div>
    </div>
  );
}

/** 거래명세서 모달 (A4 비율, 인쇄 + 이미지 저장) */
function ReceiptModal({ sale, items, customerType, onClose }: {
  sale: OfflineSale;
  items: OfflineSaleItem[];
  customerType?: string;
  onClose: () => void;
}) {
  const receiptRef = useRef<HTMLDivElement>(null);
  const [saving, setSaving] = useState(false);
  const PAYMENT_METHOD: Record<string, string> = { card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합' };
  const { data: allProducts = [] } = useProducts();
  const priceGroups = usePriceGroups();

  // 품목명 편집용 로컬 state — 납품명 우선 적용
  const [editNames, setEditNames] = useState<string[]>(() => items.map(i => i.product_name));

  // allProducts 로드 후 납품명 자동 적용 (최초 1회)
  const [namesApplied, setNamesApplied] = useState(false);
  if (!namesApplied && allProducts.length > 0 && customerType) {
    const groupKey = Object.entries(priceGroups).find(
      ([, def]) => def.customerTypes.includes(customerType)
    )?.[0];
    if (groupKey) {
      const newNames = items.map((item) => {
        const product = allProducts.find(p => p.id === item.product_id);
        const displayName = product?.price_groups?.[groupKey]?.display_name;
        return displayName || item.product_name;
      });
      setEditNames(newNames);
    }
    setNamesApplied(true);
  }

  const handlePrint = () => {
    const el = receiptRef.current;
    if (!el) return;
    // input value를 텍스트로 치환한 HTML 생성
    const clone = el.cloneNode(true) as HTMLElement;
    clone.querySelectorAll('input').forEach((input) => {
      const span = document.createElement('span');
      span.textContent = input.value;
      span.style.cssText = input.style.cssText;
      input.replaceWith(span);
    });
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.write(`
      <html><head><title>거래명세서 — ${sale.sale_number}</title>
      <style>
        body { font-family: 'Noto Sans KR', sans-serif; font-size: 12px; color: #000; padding: 20mm; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 6px 8px; text-align: left; }
        thead th { border-bottom: 2px solid #000; }
        tbody td { border-bottom: 1px solid #ddd; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .bold { font-weight: bold; }
        .total-row td { border-top: 2px solid #000; font-weight: bold; }
      </style></head><body>${clone.innerHTML}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  const handleSaveImage = async () => {
    const el = receiptRef.current;
    if (!el) return;
    setSaving(true);
    try {
      // input을 임시로 텍스트로 치환 후 캡처
      const inputs = el.querySelectorAll('input');
      const originals: { input: HTMLInputElement; parent: Node; next: Node | null }[] = [];
      inputs.forEach((input) => {
        const span = document.createElement('span');
        span.textContent = input.value;
        span.style.fontSize = '13px';
        originals.push({ input, parent: input.parentNode!, next: input.nextSibling });
        input.replaceWith(span);
      });
      const html2canvas = (await import('html2canvas')).default;
      const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff' });
      // 원래 input 복원
      originals.forEach(({ input, parent, next }) => {
        const span = next ? (next as Element).previousSibling : parent.lastChild;
        if (span) parent.replaceChild(input, span);
      });
      const link = document.createElement('a');
      link.download = `거래명세서_${sale.sale_number}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();
    } catch (e) {
      console.error('이미지 저장 실패:', e);
    } finally {
      setSaving(false);
    }
  };

  const handleCopyImage = async () => {
    const el = receiptRef.current;
    if (!el) return;
    setSaving(true);
    try {
      const inputs = el.querySelectorAll('input');
      const originals: { input: HTMLInputElement; parent: Node; next: Node | null }[] = [];
      inputs.forEach((input) => {
        const span = document.createElement('span');
        span.textContent = input.value;
        span.style.fontSize = '13px';
        originals.push({ input, parent: input.parentNode!, next: input.nextSibling });
        input.replaceWith(span);
      });
      const html2canvas = (await import('html2canvas')).default;
      const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff' });
      originals.forEach(({ input, parent, next }) => {
        const span = next ? (next as Element).previousSibling : parent.lastChild;
        if (span) parent.replaceChild(input, span);
      });
      canvas.toBlob(async (blob) => {
        if (!blob) return;
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
        toast.success('이미지가 클립보드에 복사되었습니다');
      }, 'image/png');
    } catch (e) {
      console.error('이미지 복사 실패:', e);
      toast.error('이미지 복사 실패');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '595px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* 상단 버튼 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">거래명세서</h3>
          <div className="flex items-center gap-2">
            <button
              onClick={handleCopyImage}
              disabled={saving}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-100 text-xs text-neutral-700 hover:bg-neutral-200 transition disabled:opacity-50"
            >
              <Copy size={12} />
              이미지 복사
            </button>
            <button
              onClick={handleSaveImage}
              disabled={saving}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-100 text-xs text-neutral-700 hover:bg-neutral-200 transition disabled:opacity-50"
            >
              <Download size={12} />
              이미지 저장
            </button>
            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition"
            >
              <Printer size={12} />
              인쇄
            </button>
            <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
          </div>
        </div>

        {/* A4 비율 콘텐츠 */}
        <div className="overflow-y-auto flex-1 p-6" ref={receiptRef}>
          {/* 헤더 */}
          <h1 style={{ fontSize: '20px', fontWeight: 'bold', textAlign: 'center', marginBottom: '24px' }}>
            거 래 명 세 서
          </h1>

          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px', fontSize: '13px' }}>
            <div>
              <p style={{ fontWeight: 'bold', fontSize: '15px' }}>MAMORU (마모루)</p>
              <p style={{ color: '#666' }}>미용가위 전문 브랜드</p>
              <p style={{ color: '#666' }}>서울특별시 구로구 부광로 88 SKV1, B동 311호</p>
              <p style={{ color: '#666' }}>TEL: 02-6326-0426</p>
            </div>
            <div style={{ textAlign: 'right' }}>
              <p>판매번호: {sale.sale_number}</p>
              <p>판매일: {sale.sale_date}</p>
              <p>발행일: {new Date().toISOString().slice(0, 10)}</p>
            </div>
          </div>

          {/* 고객 */}
          <div style={{ border: '1px solid #ddd', borderRadius: '4px', padding: '10px 14px', marginBottom: '16px', fontSize: '13px' }}>
            <span style={{ color: '#888' }}>고객명: </span>
            <strong>{sale.customer_name}</strong><span style={{ color: '#888', fontWeight: 'normal' }}> 님</span>
            {sale.customer_phone && (
              <span style={{ marginLeft: '24px' }}><span style={{ color: '#888' }}>연락처: </span>{sale.customer_phone}</span>
            )}
          </div>

          {/* 품목 */}
          <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '16px', fontSize: '13px' }}>
            <thead>
              <tr style={{ borderTop: '2px solid #000', borderBottom: '2px solid #000' }}>
                <th style={{ padding: '8px', textAlign: 'left' }}>품명</th>
                <th style={{ padding: '8px', textAlign: 'center', width: '60px' }}>수량</th>
                <th style={{ padding: '8px', textAlign: 'right', width: '100px' }}>단가</th>
                <th style={{ padding: '8px', textAlign: 'right', width: '120px' }}>금액</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, i) => (
                <tr key={i} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '4px 8px' }}>
                    <input
                      value={editNames[i] || ''}
                      onChange={(e) => setEditNames(prev => { const next = [...prev]; next[i] = e.target.value; return next; })}
                      style={{ width: '100%', border: 'none', borderBottom: '1px dashed #ccc', outline: 'none', fontSize: '13px', padding: '4px 0', background: 'transparent' }}
                    />
                  </td>
                  <td style={{ padding: '8px', textAlign: 'center' }}>{item.quantity}</td>
                  <td style={{ padding: '8px', textAlign: 'right' }}>{formatKRW(item.unit_price)}</td>
                  <td style={{ padding: '8px', textAlign: 'right', fontWeight: '600' }}>{formatKRW(item.total_price)}</td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* 합계 */}
          <div style={{ borderTop: '2px solid #000', paddingTop: '12px', fontSize: '13px' }}>
            {(sale.discount_amount || 0) > 0 && (
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                <span>할인</span><span>-{formatKRW(sale.discount_amount)}</span>
              </div>
            )}
            {(sale.supply_amount || 0) > 0 && (
              <>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px', color: '#666' }}>
                  <span>공급가액</span><span>{formatKRW(sale.supply_amount)}</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px', color: '#666' }}>
                  <span>부가세</span><span>{formatKRW(sale.vat_amount)}</span>
                </div>
              </>
            )}
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '16px', fontWeight: 'bold', paddingTop: '8px', borderTop: '1px solid #ddd' }}>
              <span>합계</span>
              <span>{formatKRW(sale.total_amount - (sale.discount_amount || 0))}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '4px', fontSize: '12px', color: '#888' }}>
              <span>결제방법</span>
              <span>{PAYMENT_METHOD[sale.payment_method] || sale.payment_method}</span>
            </div>
          </div>

          {/* 공급자 / 공급받는자 */}
          <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '40px', fontSize: '11px', color: '#666', gap: '24px' }}>
            <div style={{ flex: 1, border: '1px solid #ddd', borderRadius: '4px', padding: '10px' }}>
              <p style={{ fontWeight: 'bold', color: '#333', marginBottom: '4px' }}>공급자</p>
              <p>MAMORU (마모루)</p>
              <p>서울특별시 구로구 부광로 88 SKV1, B동 311호</p>
              <p>TEL: 02-6326-0426</p>
            </div>
            <div style={{ flex: 1, border: '1px solid #ddd', borderRadius: '4px', padding: '10px' }}>
              <p style={{ fontWeight: 'bold', color: '#333', marginBottom: '4px' }}>공급받는자</p>
              <p>{sale.customer_name} <span style={{ color: '#888', fontWeight: 'normal' }}>님</span></p>
              {sale.customer_phone && <p>TEL: {sale.customer_phone}</p>}
            </div>
          </div>

          {/* 하단 로고 */}
          <div style={{ textAlign: 'center', marginTop: '32px', opacity: 0.3 }}>
            <p style={{ fontSize: '14px', fontWeight: 'bold', letterSpacing: '4px' }}>MAMORU</p>
          </div>
        </div>
      </div>
    </div>
  );
}


