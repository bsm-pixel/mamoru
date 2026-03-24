'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Consultation } from '@/lib/supabase/types';

// ============================================
// 허브 대시보드 통계
// ============================================

/** 허브 대시보드 통계 타입 */
interface HubStatsResult {
  orders: {
    payDone: number;
    preparing: number;
    shipping: number;
    delivered: number;
    weekAmount: number;
    monthAmount: number;
  };
  consultations: {
    newIntake: number;
    confirmed: number;
    needAction: number;
  };
  repairs: {
    intakeNew: number;
    pendingInbound: number;
    workingCount: number;
    workingQty: number;
    weekRepairTotal: number;
    weekRepairMamoru: number;
    weekRepairOther: number;
  };
}

/** R3+P12: 허브 대시보드 — RPC 1회 호출로 통합 (fallback: 기존 14개 쿼리) */
export function useHubStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['hub-stats'],
    staleTime: 30_000, // 30초 — RPC 1회면 충분
    queryFn: async (): Promise<HubStatsResult> => {
      // RPC 호출 시도 (018_hub_stats_rpc.sql 배포 후 동작)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: rpcData, error: rpcError } = await (supabase as any).rpc('get_hub_stats');

      if (!rpcError && rpcData) {
        const d = rpcData as HubStatsResult;
        return {
          orders: {
            payDone: d.orders?.payDone ?? 0,
            preparing: d.orders?.preparing ?? 0,
            shipping: d.orders?.shipping ?? 0,
            delivered: d.orders?.delivered ?? 0,
            weekAmount: d.orders?.weekAmount ?? 0,
            monthAmount: d.orders?.monthAmount ?? 0,
          },
          consultations: {
            newIntake: d.consultations?.newIntake ?? 0,
            confirmed: d.consultations?.confirmed ?? 0,
            needAction: d.consultations?.needAction ?? 0,
          },
          repairs: {
            intakeNew: d.repairs?.intakeNew ?? 0,
            pendingInbound: Math.max(0, d.repairs?.pendingInbound ?? 0),
            workingCount: d.repairs?.workingCount ?? 0,
            workingQty: d.repairs?.workingQty ?? 0,
            weekRepairTotal: d.repairs?.weekRepairTotal ?? 0,
            weekRepairMamoru: d.repairs?.weekRepairMamoru ?? 0,
            weekRepairOther: d.repairs?.weekRepairOther ?? 0,
          },
        };
      }

      // Fallback: RPC 미배포 시 기존 14개 쿼리
      const now = new Date();
      const dow = now.getDay();
      const monday = new Date(now);
      monday.setDate(now.getDate() - ((dow + 6) % 7));
      monday.setHours(0, 0, 0, 0);
      const monISO = monday.toISOString();
      const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
      const monthISO = monthStart.toISOString();

      const [
        payDone, preparing, shipping, delivered,
        weekOrders, monthOrders,
        newConsult, confirmedConsult, needAction,
        intakeNew, repairPending, repairWorking,
        weekRepairs,
      ] = await Promise.all([
        supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'pay_done'),
        supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'preparing'),
        supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'shipping'),
        supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'delivered'),
        supabase.from('orders').select('paid_amount').gte('ordered_at', monISO).not('status', 'in', '("cancelled","refunded")'),
        supabase.from('orders').select('paid_amount').gte('ordered_at', monthISO).not('status', 'in', '("cancelled","refunded")'),
        supabase.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'pending_admin'),
        supabase.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'confirmed'),
        supabase.from('consultations').select('*', { count: 'exact', head: true })
          .in('status', ['reschedule_requested', 'change_requested', 'pending_admin'])
          .in('consultation_type', ['field_request', 'talk_consult']),
        supabase.from('repairs').select('*', { count: 'exact', head: true })
          .eq('status', 'intake').is('confirmed_at', null),
        supabase.from('repairs').select('*', { count: 'exact', head: true })
          .in('status', ['intake', 'pickup_scheduled']),
        supabase.from('repairs').select('qty_mamoru, qty_other')
          .in('status', ['cost_notified', 'repairing', 'ready_to_ship']),
        supabase.from('repairs').select('qty_mamoru, qty_other')
          .in('status', ['shipped', 'delivered', 'completed'])
          .gte('shipped_at', monISO),
      ]);

      const sumAmount = (rows: { paid_amount?: number }[]) =>
        (rows || []).reduce((s, r) => s + (r.paid_amount || 0), 0);
      const weekAmount = sumAmount(weekOrders.data || []);
      const monthAmount = sumAmount(monthOrders.data || []);

      const workingRows = (repairWorking.data || []) as { qty_mamoru: number; qty_other: number }[];
      const workingQty = workingRows.reduce((s, r) => s + (r.qty_mamoru || 0) + (r.qty_other || 0), 0);
      const workingCount = workingRows.length;

      const weekRepairRows = (weekRepairs.data || []) as { qty_mamoru: number; qty_other: number }[];
      const weekRepairMamoru = weekRepairRows.reduce((s, r) => s + (r.qty_mamoru || 0), 0);
      const weekRepairOther = weekRepairRows.reduce((s, r) => s + (r.qty_other || 0), 0);

      const pendingInbound = (repairPending.count || 0) - (intakeNew.count || 0);

      return {
        orders: {
          payDone: payDone.count || 0,
          preparing: preparing.count || 0,
          shipping: shipping.count || 0,
          delivered: delivered.count || 0,
          weekAmount,
          monthAmount,
        },
        consultations: {
          newIntake: newConsult.count || 0,
          confirmed: confirmedConsult.count || 0,
          needAction: needAction.count || 0,
        },
        repairs: {
          intakeNew: intakeNew.count || 0,
          pendingInbound: Math.max(0, pendingInbound),
          workingCount,
          workingQty,
          weekRepairTotal: weekRepairRows.length,
          weekRepairMamoru,
          weekRepairOther,
        },
      };
    },
  });
}

// ============================================
// 주문 전용 대시보드 통계
// ============================================

export function useOrderDashboardStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['order-dashboard-stats'],
    staleTime: 30_000,
    queryFn: async () => {
      const todayISO = new Date(new Date().setHours(0, 0, 0, 0)).toISOString();

      const [payDone, preparing, shipping, delivered, todayOrders] =
        await Promise.all([
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'pay_done'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'preparing'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'shipping'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'delivered'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).gte('ordered_at', todayISO),
        ]);

      return {
        payDone: payDone.count || 0,
        preparing: preparing.count || 0,
        shipping: shipping.count || 0,
        delivered: delivered.count || 0,
        todayOrders: todayOrders.count || 0,
        pipeline: [
          { label: '결제완료', count: payDone.count || 0, status: 'pay_done' },
          { label: '준비중', count: preparing.count || 0, status: 'preparing' },
          { label: '배송중', count: shipping.count || 0, status: 'shipping' },
          { label: '배송완료', count: delivered.count || 0, status: 'delivered' },
        ],
      };
    },
  });
}

// ============================================
// 상담 전용 대시보드 통계
// ============================================

export function useConsultationDashboardStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['consultation-dashboard-stats'],
    staleTime: 30_000,
    queryFn: async () => {
      const now = new Date();
      const todayStr = now.toISOString().slice(0, 10);
      // R2: 6시간 기준
      const sixHoursAgo = new Date(now.getTime() - 6 * 60 * 60 * 1000).toISOString();
      // 1달 전
      const oneMonthAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000).toISOString();

      const [
        newIntake,
        inProgress,
        completedMonth,
        todaySchedule,
      ] = await Promise.all([
        // R2: 신규접수 (6시간 이내 + 미처리)
        supabase
          .from('consultations')
          .select('*', { count: 'exact', head: true })
          .gte('received_at', sixHoursAgo)
          .in('status', ['pending_admin', 'confirmed']),
        // R2: 진행중 (6시간 이후 + 미완료/미취소)
        supabase
          .from('consultations')
          .select('*', { count: 'exact', head: true })
          .lt('received_at', sixHoursAgo)
          .not('status', 'in', '("completed","cancelled")'),
        // R2: 상담완료 (1달 이내)
        supabase
          .from('consultations')
          .select('*', { count: 'exact', head: true })
          .eq('status', 'completed')
          .gte('updated_at', oneMonthAgo),
        // 오늘 일정 리스트 (full rows)
        supabase
          .from('consultations')
          .select('*')
          .eq('visit_date', todayStr)
          .in('status', ['confirmed', 'assigned', 'in_progress', 'pending_admin'])
          .order('visit_time', { ascending: true })
          .limit(10),
      ]);

      return {
        newIntake: newIntake.count || 0,
        inProgress: inProgress.count || 0,
        completedMonth: completedMonth.count || 0,
        todaySchedule: (todaySchedule.data || []) as Consultation[],
      };
    },
  });
}

// ============================================
// 복원수리 전용 대시보드 통계
// ============================================

export function useRepairDashboardStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['repair-dashboard-stats'],
    staleTime: 30_000,
    queryFn: async () => {
      // 3일 전 기준
      const threeDaysAgo = new Date();
      threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
      const threeDaysAgoISO = threeDaysAgo.toISOString();

      const statuses = [
        'intake', 'pickup_scheduled',
        'cost_notified', 'repairing',
        'ready_to_ship', 'shipped', 'delivered',
      ] as const;

      const [counts, staleCount, unpaidCount, intakeNewCount] = await Promise.all([
        // 상태별 count 병렬
        Promise.all(
          statuses.map((s) =>
            supabase
              .from('repairs')
              .select('*', { count: 'exact', head: true })
              .eq('status', s)
              .then((r) => ({ status: s, count: r.count || 0 }))
          )
        ),
        // 경과일 3일 이상 미처리 (intake, cost_notified)
        supabase
          .from('repairs')
          .select('*', { count: 'exact', head: true })
          .in('status', ['intake', 'cost_notified'])
          .lt('updated_at', threeDaysAgoISO),
        // 미입금 건수 (cost_notified 이후 + paid_at IS NULL)
        supabase
          .from('repairs')
          .select('*', { count: 'exact', head: true })
          .in('status', ['cost_notified', 'repairing', 'ready_to_ship', 'shipped'])
          .is('paid_at', null),
        // R1: 신규접수 (intake + confirmed_at IS NULL)
        supabase
          .from('repairs')
          .select('*', { count: 'exact', head: true })
          .eq('status', 'intake')
          .is('confirmed_at', null),
      ]);

      const byStatus = Object.fromEntries(counts.map((c) => [c.status, c.count])) as Record<string, number>;

      // 진행대기 = intake+pickup_scheduled 중 confirmed_at 있는 것 (입고 전 전체)
      const pendingInbound = (byStatus.intake || 0) + (byStatus.pickup_scheduled || 0) - (intakeNewCount.count || 0);
      // 작업중 = cost_notified + repairing + ready_to_ship
      const workingCount = (byStatus.cost_notified || 0) + (byStatus.repairing || 0) + (byStatus.ready_to_ship || 0);

      return {
        byStatus,
        staleCount: staleCount.count || 0,
        unpaidCount: unpaidCount.count || 0,
        // R1: 탭별 카운트 (대시보드 카드용)
        intakeNew: intakeNewCount.count || 0,
        pendingInbound,
        workingCount,
        pipeline: [
          { label: '신규접수', count: intakeNewCount.count || 0, status: 'intake' },
          { label: '입고대기', count: pendingInbound, status: 'pickup_scheduled' },
          { label: '진행중', count: workingCount, status: 'repairing' },
          { label: '출고대기', count: byStatus.ready_to_ship || 0, status: 'ready_to_ship' },
          { label: '출고완료', count: byStatus.shipped || 0, status: 'shipped' },
          { label: '배송완료', count: byStatus.delivered || 0, status: 'delivered' },
        ],
      };
    },
  });
}

// ============================================
// 미수금 경고 (outstanding_balance > 0인 고객)
// ============================================

export function useOutstandingAlert() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['outstanding-alert'],
    staleTime: 60_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any)
        .from('customers')
        .select('id, name, phone, outstanding_balance')
        .gt('outstanding_balance', 0)
        .order('outstanding_balance', { ascending: false })
        .limit(10);
      if (error) throw error;
      return (data || []) as Array<{
        id: string;
        name: string;
        phone: string | null;
        outstanding_balance: number;
      }>;
    },
  });
}
