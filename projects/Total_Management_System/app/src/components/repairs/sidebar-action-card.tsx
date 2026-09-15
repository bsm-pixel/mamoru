'use client';

import { useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { DeliveryTracker } from '@/components/orders/delivery-tracker';
import { ActionNote, MoreActions, DangerZone, DangerLink, SubtleButton } from '@/components/ui/action-section';
import { isAlpsTrackable, courierLabel } from '@/lib/shipping/couriers';
import {
  useUpdateRepairStatus,
  useUpdateRepairFields,
  useShipRepair,
  useCancelShipment,
  useRecallRepairPickup,
  useReworkRepair,
  useSendRepairNotification,
  useDeleteRepair,
} from '@/hooks/use-repairs';
import { invalidateFinancialQueries } from '@/lib/query/invalidate-keys';
import { getFilteredRepairTransitions, REPAIR_ACTION_LABEL, getRepairDisplayLabel } from '@/lib/repair/transitions';
import { formatKRW, formatDateTime } from '@/lib/utils/format';
import type { Repair, RepairStatus } from '@/lib/supabase/types';
import { Package, Truck, Send, CheckCircle, CreditCard } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { MergedShipModal } from './merged-ship-modal';

interface SidebarActionCardProps {
  repair: Repair;
}

type ConfirmAction =
  | 'cost_notice' | 'mark_paid' | 'mark_shipped' | 'cancel_shipment'
  | 'cancel_repair' | 'delete_repair' | 'visit_checkout'
  | 'direct_pickup' | 'recall_pickup' | 'rework' | 'mark_delivered'
  | null;
type PayMethod = 'transfer' | 'card' | 'cash';
const PAY_METHODS: { id: PayMethod; label: string }[] = [
  { id: 'transfer', label: '이체' },
  { id: 'card', label: '카드' },
  { id: 'cash', label: '현금' },
];

/**
 * 복원수리 사이드바 — 표준 5슬롯 액션 (components/ui/action-section.tsx)
 *
 * 2026-09-15 재설계(사장님 지적: 배치가 난잡하다):
 *   전엔 비용 카드 안에 입금확인·비용안내·상태전이·취소·삭제가 섞여 있었고,
 *   출고 카드에는 출고완료·송장취소·수동배송완료·재수거·재작업·직접수령이 평면 나열돼
 *   "지금 눌러야 할 것"이 무엇인지 화면만 봐선 알 수 없었다.
 *   재수거(파랑)·재작업(주황) 채움 버튼은 브랜드 모노크롬도 깨고 있었다.
 *
 *   이제 카드 역할을 나눈다: 비용=정보 / 출고=정보 / 액션=아래 5슬롯 고정 순서
 *     ① 지금 상태 한 줄 ② 주 액션(검은 버튼 1개) ③ 결제 ④ 접힘 ⑤ 파괴적 액션
 *   ⚠️ 슬롯 순서를 바꾸지 말 것 — 납품·판매·주문 상세와 같은 순서로 통일돼 있다.
 */
export function SidebarActionCard({ repair: r }: SidebarActionCardProps) {
  const queryClient = useQueryClient();
  const updateStatus = useUpdateRepairStatus();
  const updateFields = useUpdateRepairFields();
  const shipRepair = useShipRepair();
  const cancelShipment = useCancelShipment();
  const recallPickup = useRecallRepairPickup();
  const reworkRepair = useReworkRepair();
  const sendNotify = useSendRepairNotification();
  const deleteRepair = useDeleteRepair();
  const [confirmAction, setConfirmAction] = useState<ConfirmAction>(null);
  const [mergedShipOpen, setMergedShipOpen] = useState(false);
  const [payMethod, setPayMethod] = useState<PayMethod>('transfer'); // 120: 결제수단
  const [visitPaid, setVisitPaid] = useState(true); // 직접방문 현장결제 시 입금완료 처리 여부
  const isDirectVisit = r.proceed_type === '직접방문';

  const currentStatus = r.status as RepairStatus;
  const proceedType = r.proceed_type;
  const filtered = getFilteredRepairTransitions(currentStatus, proceedType);
  const busy = updateStatus.isPending;
  const isTerminal = currentStatus === 'completed' || currentStatus === 'cancelled';

  // 비용안내 발송 가능 상태
  const canSendCostNotice = ['intake', 'pickup_scheduled', 'picked_up', 'inspecting', 'cost_notified'].includes(currentStatus);
  const isCostResend = currentStatus === 'cost_notified';

  // 입금확인 독립 버튼 표시 조건 — 미입금이면 배송완료/완료 이후에도 가능
  //   (ALPS 자동추적으로 shipped→delivered→completed 전이 후 입금확인 길이 막히던 버그 fix 2026-06-11)
  const canMarkPaid = !r.paid_at && ['cost_notified', 'repairing', 'ready_to_ship', 'shipped', 'delivered', 'completed'].includes(currentStatus);

  // 출고완료 버튼 조건 (ready_to_ship 상태 + 송장 있음)
  const canMarkShipped = currentStatus === 'ready_to_ship' && !!r.invoice_number;

  // 합포장 출고 버튼 조건 (송장 없을 때 — 다른 주문 송장에 합쳐 발송한 케이스)
  const canMergedShip = currentStatus === 'ready_to_ship' && !r.invoice_number;

  // 119: 택배 없이 매장에서 직접 전달(직접수령) → 바로 완료 처리. 작업중/출고대기 + 송장 없을 때만
  const canDirectPickup = ['repairing', 'ready_to_ship'].includes(currentStatus) && !r.invoice_number;

  // 정밀 재점검 재수거 (출고된 건 회수 — 롯데 반품 API 02). 알림톡 없음, 취소는 ALPS 수동
  const canRecall = ['shipped', 'delivered', 'completed'].includes(currentStatus);

  // 🔒 기존 동작 보존: 출고 섹션이 cost_notified 이상에서 송장 생성을 열어줬다(입금확인 후 바로 발급).
  //    ready_to_ship 은 주 액션이라 여기서 제외 — 나머지 단계에서는 보조 버튼으로 제공한다.
  const canBookEarly = !r.invoice_number && ['cost_notified', 'repairing'].includes(currentStatus);

  // 상태 전이 버튼 — 별도 전용 버튼이 있는 전이는 제외. 실제로는 항상 0~1개다
  const visibleTransitions = filtered.filter(
    (s) => !['cancelled', 'cost_notified', 'shipped', 'repairing', 'delivered'].includes(s),
  );
  const primaryTransition = visibleTransitions[0];

  // 접힘(④)에 들어갈 게 하나라도 있는가
  const hasMore =
    visibleTransitions.length > 1 ||
    canMergedShip || canDirectPickup || canRecall ||
    (!!r.invoice_number && ['ready_to_ship', 'shipped'].includes(currentStatus)) ||
    currentStatus === 'shipped' ||
    (!!primaryTransition && canSendCostNotice && !isDirectVisit && visibleTransitions.length > 1);

  const handleSendCostNotice = async () => {
    const isFree = r.total_amount === 0; // 무상 처리 여부

    // 비용안내 → 자동으로 repairing 전환 (작업시작 별도 클릭 불필요)
    await updateStatus.mutateAsync({
      id: r.id,
      status: 'cost_notified',
      service_cost: r.service_cost,
      shipping_fee: r.shipping_fee,
      total_amount: r.total_amount,
      note: isFree ? '무상 처리 비용 안내' : `비용 안내: ${formatKRW(r.total_amount)}`,
    });

    // 0원이면 자동 입금완료 처리
    if (isFree && !r.paid_at) {
      await updateFields.mutateAsync({
        id: r.id,
        paid_at: new Date().toISOString(), // 무상(0원) — 서버가 총액 0 을 보고 입금확인 알림톡 생략(비용안내와 중복 방지)
      });
    }

    // 비용안내 발송 후 자동으로 수리중 전환
    await updateStatus.mutateAsync({
      id: r.id,
      status: 'repairing',
      note: '비용안내 후 자동 작업시작',
    });
    sendNotify.mutate({
      repairId: r.id,
      template: 'as_cost_notice',
      extraData: {
        as_amount: String(r.service_cost),
        shipping_amount: String(r.shipping_fee),
        total_amount: String(r.total_amount),
      },
    });
  };

  // 입금확인 처리 (paid_at 설정 — 상태 변경 없음) + 결제수단 기록
  const handleMarkPaid = () => {
    updateFields.mutate({
      id: r.id,
      paid_at: new Date().toISOString(),
      payment_method: payMethod, // 120
    });
  };

  // 직접방문 현장결제: 비용확정→작업시작(알림톡 없음) + 결제수단 + 입금(옵션) 한 번에
  const handleVisitCheckout = async () => {
    if (currentStatus === 'intake') {
      await updateStatus.mutateAsync({
        id: r.id,
        status: 'cost_notified',
        service_cost: r.service_cost,
        shipping_fee: r.shipping_fee,
        total_amount: r.total_amount,
        note: `방문 현장 비용확정: ${formatKRW(r.total_amount)}`,
      });
    }
    await updateStatus.mutateAsync({ id: r.id, status: 'repairing', note: '방문 현장 작업시작' });
    await updateFields.mutateAsync({
      id: r.id,
      payment_method: payMethod,
      // 현장결제 — 고객이 매장에서 직접 결제해 입금확인 알림톡은 서버가 생략(paid_on_site)
      ...(visitPaid && !r.paid_at ? { paid_at: new Date().toISOString(), paid_on_site: true } : {}),
    });
    setConfirmAction(null);
  };

  // 출고완료 처리 (ready_to_ship → shipped + shipped_at 설정)
  const handleMarkShipped = () => {
    updateStatus.mutate({
      id: r.id,
      status: 'shipped',
      shipped_at: new Date().toISOString(),
      note: '출고완료',
    });
  };

  // 합포장 출고 성공 시 캐시 갱신 (모달 자체 API 호출 → 부모에서 invalidate)
  const handleMergedShipSuccess = () => {
    queryClient.invalidateQueries({ queryKey: ['repair', r.id] });
    queryClient.invalidateQueries({ queryKey: ['repair-tabs'] });
    invalidateFinancialQueries(queryClient);
  };

  // ① 상태 한 줄 — 지금 어디에 있고 다음에 뭐가 자동으로 되는지
  function renderNote() {
    const title = getRepairDisplayLabel(currentStatus, proceedType);
    if (currentStatus === 'shipped') {
      return (
        <ActionNote tone="done" title="출고완료" meta={r.shipped_at ? formatDateTime(r.shipped_at) : undefined}>
          {r.invoice_number && isAlpsTrackable(r.courier_name)
            ? '인수자등록이 감지되면 배송완료로 바뀝니다 (1시간마다 확인)'
            : '자동 배송추적은 되지 않습니다. 받으셨으면 접힘에서 배송완료를 눌러주세요.'}
        </ActionNote>
      );
    }
    if (currentStatus === 'delivered') {
      return <ActionNote tone="done" title="배송완료">고객 수령이 확인되면 완료로 넘겨주세요.</ActionNote>;
    }
    if (currentStatus === 'completed') {
      return <ActionNote tone="done" title="완료">끝난 건입니다. 재점검이 필요하면 아래에서 재수거를 접수하세요.</ActionNote>;
    }
    if (currentStatus === 'ready_to_ship') {
      return (
        <ActionNote
          tone="wait"
          title="출고대기"
          meta={r.invoice_number ? `${courierLabel(r.courier_name)} ${r.invoice_number}` : '송장 없음'}
        >
          {r.invoice_number
            ? (isAlpsTrackable(r.courier_name)
                ? '롯데 기사님이 수거하면 자동으로 출고완료 처리됩니다 (1시간마다 확인)'
                : '자동 감지가 안 되는 택배사입니다. 보내셨으면 [출고완료]를 눌러주세요.')
            : '송장을 만들거나, 판매건에 합포장하거나, 매장에서 직접 전달할 수 있습니다.'}
        </ActionNote>
      );
    }
    if (currentStatus === 'repairing') {
      return <ActionNote tone="wait" title="작업중">수리가 끝나면 출고대기로 넘겨주세요.</ActionNote>;
    }
    if (currentStatus === 'cost_notified') {
      return <ActionNote tone="wait" title="작업중" meta="비용안내 발송됨">입금이 확인되면 아래 [입금확인]을 눌러주세요.</ActionNote>;
    }
    if (currentStatus === 'cancelled') {
      return <ActionNote tone="muted" title="취소됨" />;
    }
    return (
      <ActionNote tone="wait" title={title}>
        {isDirectVisit
          ? '고객이 매장에 오시면 현장에서 비용을 확정하고 결제까지 한 번에 처리합니다.'
          : proceedType === '방문수거' && currentStatus === 'intake'
            ? '고객집 방문수거를 먼저 접수해주세요.'
            : '가위가 도착하면 검수 후 비용안내를 보내주세요.'}
      </ActionNote>
    );
  }

  return (
    <>
      {/* ── 비용 (정보만) ── */}
      <Card>
        <CardHeader>
          <CardTitle>비용</CardTitle>
        </CardHeader>
        <div className="text-sm space-y-1">
          <div className="flex justify-between">
            <span className="text-neutral-500">수리비</span>
            <span>{formatKRW(r.service_cost)}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">수거비</span>
            <span>{formatKRW(r.shipping_fee)}</span>
          </div>
          <div className="flex justify-between font-bold border-t border-neutral-100 pt-1 mt-1">
            <span>합계</span>
            <span className="text-terracotta-deep">{formatKRW(r.total_amount)}</span>
          </div>
        </div>

        {/* 입금 상태 — 처리 버튼은 액션 카드(③)로 옮겼다 */}
        {r.paid_at && (
          <div className="mt-3 pt-3 border-t border-neutral-100">
            <div className="flex items-center gap-1.5 text-sm text-success font-medium">
              <CheckCircle size={14} />
              입금완료
              <span className="text-xs text-neutral-400 font-normal ml-auto">
                {formatDateTime(r.paid_at)}
              </span>
            </div>
          </div>
        )}
      </Card>

      {/* ── 출고 (정보만) — 송장이 있을 때만 ── */}
      {r.invoice_number && (
        <Card>
          <CardHeader>
            <CardTitle>
              <Truck size={16} className="inline mr-1.5" />
              출고
            </CardTitle>
          </CardHeader>
          <div className="space-y-2">
            <div className="flex items-center gap-2 text-sm">
              <Package size={14} className="text-success" />
              <span className="font-mono font-medium">{r.invoice_number}</span>
            </div>
            <p className="text-xs text-neutral-500">
              {courierLabel(r.courier_name)}
              {!isAlpsTrackable(r.courier_name) && ' · 자동 추적 미지원'}
            </p>
            {r.shipped_at && (
              <p className="text-xs text-neutral-400">발송: {formatDateTime(r.shipped_at)}</p>
            )}
            {/* 라이브 배송추적 — 롯데 송장만 조회된다 */}
            {isAlpsTrackable(r.courier_name) && <DeliveryTracker invNo={r.invoice_number} />}
            {r.recall_invoice_number && (
              <div className="text-xs text-neutral-600 bg-stone-50 rounded-lg px-2.5 py-2">
                ↩ 재수거 송장 <b className="font-mono">{r.recall_invoice_number}</b>
                {r.recall_booked_at && <span className="text-neutral-400 ml-1">({formatDateTime(r.recall_booked_at)})</span>}
                <p className="text-[11px] text-neutral-400 mt-0.5">접수 후 취소는 ALPS에서 수동입니다.</p>
              </div>
            )}
          </div>
        </Card>
      )}

      {/* ── 액션 — 표준 5슬롯 ── */}
      {currentStatus !== 'cancelled' && (
        <Card>
          <h4 className="text-xs font-semibold text-neutral-500 mb-2">액션</h4>
          <div className="space-y-2">

            {/* ① 지금 상태 */}
            {renderNote()}

            {/* ② 주 액션 — 딱 하나 */}
            {canSendCostNotice && isDirectVisit ? (
              <Button
                className="w-full"
                onClick={() => setConfirmAction('visit_checkout')}
                loading={updateStatus.isPending || updateFields.isPending}
              >
                🏪 방문 확정 · 현장결제
              </Button>
            ) : primaryTransition ? (
              <Button
                className="w-full"
                disabled={busy}
                loading={updateStatus.variables?.status === primaryTransition && busy}
                onClick={() => updateStatus.mutate({ id: r.id, status: primaryTransition })}
              >
                {REPAIR_ACTION_LABEL[primaryTransition]}
              </Button>
            ) : canSendCostNotice ? (
              <Button
                className="w-full"
                onClick={() => setConfirmAction('cost_notice')}
                loading={updateStatus.isPending || sendNotify.isPending}
              >
                <Send size={14} />
                {isCostResend ? '비용 안내 재발송' : '입고 & 비용안내'}
              </Button>
            ) : canMarkShipped ? (
              <Button className="w-full" onClick={() => setConfirmAction('mark_shipped')} loading={updateStatus.isPending}>
                <Truck size={14} />
                출고완료
              </Button>
            ) : currentStatus === 'ready_to_ship' ? (
              <Button className="w-full" onClick={() => shipRepair.mutate({ id: r.id })} loading={shipRepair.isPending}>
                <Truck size={14} />
                송장 생성
              </Button>
            ) : null}

            {/* 수리 도중에도 송장을 미리 뽑을 수 있다 (기존 출고 섹션이 열어주던 경로) */}
            {canBookEarly && (
              <Button variant="secondary" className="w-full" onClick={() => shipRepair.mutate({ id: r.id })} loading={shipRepair.isPending}>
                <Truck size={14} />
                송장 미리 생성
              </Button>
            )}

            {/* 방문수거 접수 단계처럼 "지금 해도 되는" 다음 행동이 하나 더 있는 경우만 흰 버튼으로 */}
            {primaryTransition && canSendCostNotice && !isDirectVisit && (
              <Button
                variant="secondary"
                className="w-full"
                onClick={() => setConfirmAction('cost_notice')}
                loading={updateStatus.isPending || sendNotify.isPending}
              >
                <Send size={14} />
                {isCostResend ? '비용 안내 재발송' : '입고 & 비용안내'}
              </Button>
            )}

            {/* ③ 결제 */}
            {canMarkPaid && (
              <Button
                variant="secondary"
                className="w-full"
                onClick={() => setConfirmAction('mark_paid')}
                loading={updateFields.isPending}
              >
                <CreditCard size={14} />
                입금확인
              </Button>
            )}

            {/* ④ 드물게 쓰는 것 — 접어둔다 */}
            {hasMore && (
              <MoreActions label={currentStatus === 'completed' ? '재점검이 필요한가요?' : '다른 처리 방법'}>
                {visibleTransitions.slice(1).map((nextStatus) => (
                  <SubtleButton
                    key={nextStatus}
                    onClick={() => updateStatus.mutate({ id: r.id, status: nextStatus })}
                    disabled={busy}
                  >
                    {REPAIR_ACTION_LABEL[nextStatus]}
                  </SubtleButton>
                ))}
                {canMergedShip && (
                  <SubtleButton onClick={() => setMergedShipOpen(true)}>
                    📦 판매건 합포장 출고
                  </SubtleButton>
                )}
                {canDirectPickup && (
                  <SubtleButton onClick={() => setConfirmAction('direct_pickup')} disabled={updateStatus.isPending}>
                    🏪 직접 수령 (매장 전달) — 택배 없이 완료
                  </SubtleButton>
                )}
                {currentStatus === 'shipped' && (
                  <SubtleButton onClick={() => setConfirmAction('mark_delivered')} disabled={updateStatus.isPending}>
                    수동 배송완료 처리
                  </SubtleButton>
                )}
                {canRecall && (
                  r.recall_invoice_number ? (
                    <SubtleButton onClick={() => setConfirmAction('rework')} disabled={reworkRepair.isPending}>
                      {reworkRepair.isPending ? '처리 중…' : '↩ 재수거품 입고 · 재작업 시작'}
                    </SubtleButton>
                  ) : (
                    <SubtleButton onClick={() => setConfirmAction('recall_pickup')} disabled={recallPickup.isPending}>
                      {recallPickup.isPending ? '접수 중…' : '🚚 정밀 재점검 재수거 접수'}
                    </SubtleButton>
                  )
                )}
                {!!r.invoice_number && ['ready_to_ship', 'shipped'].includes(currentStatus) && (
                  <DangerLink onClick={() => setConfirmAction('cancel_shipment')} disabled={cancelShipment.isPending}>
                    {cancelShipment.isPending ? '취소 중...' : '송장 취소 (기록 지우기)'}
                  </DangerLink>
                )}
              </MoreActions>
            )}

            {/* ⑤ 파괴적 액션 */}
            {!isTerminal && (
              <DangerZone>
                {filtered.includes('cancelled' as RepairStatus) && (
                  <DangerLink onClick={() => setConfirmAction('cancel_repair')} disabled={busy}>
                    복원수리 취소
                  </DangerLink>
                )}
                <DangerLink onClick={() => setConfirmAction('delete_repair')} disabled={busy}>
                  삭제
                </DangerLink>
              </DangerZone>
            )}
          </div>
        </Card>
      )}

      {/* 확인 모달들 */}
      <ConfirmModal
        open={confirmAction === 'cost_notice'}
        onClose={() => setConfirmAction(null)}
        onConfirm={handleSendCostNotice}
        title="비용 안내 발송"
        message={<>고객에게 <strong>비용 안내 알림톡</strong>이 발송됩니다.<br />금액: {formatKRW(r.total_amount)}{r.total_amount === 0 ? ' (무상 처리)' : ''}</>}
        confirmLabel="발송"
      />
      {/* 직접방문 현장결제 — 비용확정 + 결제수단 + 입금완료 한 번에 (알림톡 없음) */}
      <ConfirmModal
        open={confirmAction === 'visit_checkout'}
        onClose={() => { setConfirmAction(null); setVisitPaid(true); }}
        onConfirm={handleVisitCheckout}
        title="방문 확정 · 현장결제"
        message={
          <div className="space-y-3">
            <div className="text-sm text-neutral-600 bg-neutral-50 rounded-lg p-2.5 leading-relaxed">
              <div>수량: 마모루 {r.qty_mamoru}자루{r.qty_other > 0 ? ` · 타사 ${r.qty_other}자루` : ''}</div>
              <div className="font-semibold text-neutral-900 mt-0.5">금액: {formatKRW(r.total_amount)}{r.total_amount === 0 ? ' (무상)' : ''}</div>
              <div className="text-xs text-neutral-400 mt-1">※ 수량이 실제와 다르면 접수정보에서 먼저 조정하세요</div>
            </div>
            <div>
              <p className="text-xs text-neutral-500 mb-1.5">결제수단</p>
              <div className="flex gap-2">
                {PAY_METHODS.map((m) => (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => setPayMethod(m.id)}
                    className={`flex-1 py-1.5 rounded-lg text-sm font-semibold border transition ${payMethod === m.id ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-600 border-neutral-200 hover:bg-neutral-50'}`}
                  >
                    {m.label}
                  </button>
                ))}
              </div>
            </div>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={visitPaid}
                onChange={(e) => setVisitPaid(e.target.checked)}
                className="w-4 h-4 rounded border-neutral-300 text-terracotta focus:ring-terracotta"
              />
              <span className="text-sm text-neutral-600">입금 완료로 처리 (현장 결제)</span>
            </label>
            <p className="text-xs text-neutral-400">알림톡은 발송하지 않습니다 (고객 현장 방문).</p>
          </div>
        }
        confirmLabel="현장 처리 완료"
      />
      <ConfirmModal
        open={confirmAction === 'mark_paid'}
        onClose={() => setConfirmAction(null)}
        onConfirm={handleMarkPaid}
        title="입금 확인"
        message={
          <div className="space-y-3">
            <p>{r.name}님의 입금을 확인 처리합니다.<br />금액: {formatKRW(r.total_amount)}</p>
            <div>
              <p className="text-xs text-neutral-500 mb-1.5">결제수단</p>
              <div className="flex gap-2">
                {PAY_METHODS.map((m) => (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => setPayMethod(m.id)}
                    className={`flex-1 py-1.5 rounded-lg text-sm font-semibold border transition ${payMethod === m.id ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-600 border-neutral-200 hover:bg-neutral-50'}`}
                  >
                    {m.label}
                  </button>
                ))}
              </div>
            </div>
          </div>
        }
        confirmLabel="입금 확인 + 알림톡"
      />
      <ConfirmModal
        open={confirmAction === 'mark_shipped'}
        onClose={() => setConfirmAction(null)}
        onConfirm={handleMarkShipped}
        title="출고 완료"
        message={<>송장 {r.invoice_number}으로 출고 완료 처리합니다.<br />고객에게 <strong>출고 알림톡</strong>이 자동 발송됩니다.</>}
        confirmLabel="출고 완료"
      />
      {/* 119: 택배 없이 매장에서 직접 전달 → 바로 완료 */}
      <ConfirmModal
        open={confirmAction === 'direct_pickup'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => {
          updateStatus.mutate({
            id: r.id,
            status: 'delivered',
            delivered_at: new Date().toISOString(),
            delivery_method: 'pickup',
            note: '직접 수령 (매장 전달)',
          });
          setConfirmAction(null);
        }}
        title="직접 수령 (매장 전달)"
        message={<>{r.name}님께 매장에서 <strong>직접 전달</strong>(택배 없이)한 것으로 처리합니다.<br />배송완료 상태가 됩니다.</>}
        confirmLabel="직접 수령 완료"
      />
      <ConfirmModal
        open={confirmAction === 'mark_delivered'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => {
          updateStatus.mutate({
            id: r.id,
            status: 'delivered',
            delivered_at: new Date().toISOString(),
            note: '수동 배송완료 처리 (ALPS 추적 fallback)',
          });
          setConfirmAction(null);
        }}
        title="수동 배송완료 처리"
        message="고객이 받으신 것으로 배송완료 처리합니다. (ALPS 자동 감지가 안 될 때 사용)"
        confirmLabel="배송완료"
      />
      <ConfirmModal
        open={confirmAction === 'recall_pickup'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => { recallPickup.mutate(r.id); setConfirmAction(null); }}
        title="정밀 재점검 재수거 접수"
        message={<>고객집 방문수거로 <strong>재수거</strong>를 접수합니다.<br />접수 후 취소는 ALPS에서 수동입니다.</>}
        confirmLabel="재수거 접수"
      />
      <ConfirmModal
        open={confirmAction === 'rework'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => { reworkRepair.mutate(r.id); setConfirmAction(null); }}
        title="재수거품 입고 · 재작업 시작"
        message={<>재수거품이 입고되었습니다. <strong>수리중으로 되돌립니다.</strong><br />재출고 시 새 송장과 출고 알림톡이 발송됩니다.</>}
        confirmLabel="재작업 시작"
      />
      <MergedShipModal
        open={mergedShipOpen}
        onClose={() => setMergedShipOpen(false)}
        repairId={r.id}
        onSuccess={handleMergedShipSuccess}
      />
      <ConfirmModal
        open={confirmAction === 'cancel_shipment'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => cancelShipment.mutateAsync({ id: r.id })}
        title="송장 취소"
        message={
          <div className="space-y-2">
            <p>송장 <strong className="font-mono">{r.invoice_number}</strong> 기록을 지웁니다.</p>
            <p className="text-xs text-red-500">먼저 롯데 ALPS 에서 집하취소를 완료해주세요.</p>
          </div>
        }
        confirmLabel="송장 취소"
        variant="danger"
      />
      <ConfirmModal
        open={confirmAction === 'cancel_repair'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => updateStatus.mutateAsync({ id: r.id, status: 'cancelled', note: '관리자 취소' })}
        title="복원수리 취소"
        message="정말 이 복원수리 접수를 취소하시겠습니까?"
        confirmLabel="취소 확정"
        variant="danger"
      />
      <ConfirmModal
        open={confirmAction === 'delete_repair'}
        onClose={() => setConfirmAction(null)}
        onConfirm={async () => { await deleteRepair.mutateAsync(r.id); }}
        title="복원수리 삭제"
        message={<>이 건을 <strong>완전히 삭제</strong>합니다.<br />알림톡은 발송되지 않습니다. 복구할 수 없습니다.</>}
        confirmLabel="삭제"
        variant="danger"
      />
    </>
  );
}
