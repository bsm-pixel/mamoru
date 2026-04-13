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
    readyToShip: number;
    weekRepairTotal: number;
    weekRepairMamoru: number;
    weekRepairOther: number;
    monthRepairAmount: number;   // 복원수리 매출 (접수시스템 + 판매)
    monthRepairCount: number;
    monthRepairMamoru: { amount: number; count: number };
    monthRepairOther: { amount: number; count: number };
    monthRepairB2B: { amount: number; count: number };
  };
  sales: {
    monthCount: number;
    monthAmount: number;
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
            readyToShip: d.repairs?.readyToShip ?? 0,
            weekRepairTotal: d.repairs?.weekRepairTotal ?? 0,
            weekRepairMamoru: d.repairs?.weekRepairMamoru ?? 0,
            weekRepairOther: d.repairs?.weekRepairOther ?? 0,
            monthRepairAmount: d.repairs?.monthRepairAmount ?? 0,
            monthRepairCount: d.repairs?.monthRepairCount ?? 0,
            monthRepairMamoru: d.repairs?.monthRepairMamoru ?? { amount: 0, count: 0 },
            monthRepairOther: d.repairs?.monthRepairOther ?? { amount: 0, count: 0 },
            monthRepairB2B: d.repairs?.monthRepairB2B ?? { amount: 0, count: 0 },
          },
          sales: {
            monthCount: d.sales?.monthCount ?? 0,
            monthAmount: d.sales?.monthAmount ?? 0,
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

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const monthStartDate = monthStart.toISOString().slice(0, 10);
      const [
        payDone, preparing, shipping, delivered,
        weekOrders, monthOrders,
        newConsult, confirmedConsult, needAction,
        intakeNew, repairPending, repairWorking, readyToShipRes,
        weekRepairs,
        monthSales,
        monthRepairsPaid, monthSalesRepairItems,
        monthDeliveries,
      ] = await Promise.all([
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'pay_done'),
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'preparing'),
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'shipping'),
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'delivered'),
        db.from('orders').select('paid_amount').gte('ordered_at', monISO).not('status', 'in', '("cancelled","refunded")'),
        db.from('orders').select('paid_amount').gte('ordered_at', monthISO).not('status', 'in', '("cancelled","refunded")'),
        db.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'pending_admin'),
        db.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'confirmed'),
        db.from('consultations').select('*', { count: 'exact', head: true })
          .in('status', ['reschedule_requested', 'change_requested', 'pending_admin'])
          .in('consultation_type', ['field_request', 'talk_consult']),
        db.from('repairs').select('*', { count: 'exact', head: true })
          .eq('status', 'intake').is('confirmed_at', null),
        db.from('repairs').select('*', { count: 'exact', head: true })
          .in('status', ['intake', 'pickup_scheduled']),
        db.from('repairs').select('qty_mamoru, qty_other')
          .in('status', ['cost_notified', 'repairing']),
        db.from('repairs').select('*', { count: 'exact', head: true })
          .eq('status', 'ready_to_ship'),
        db.from('repairs').select('qty_mamoru, qty_other')
          .in('status', ['shipped', 'delivered', 'completed'])
          .gte('shipped_at', monISO),
        db.from('offline_sales').select('total_amount, discount_amount')
          .gte('sale_date', monthStartDate)
          .is('cancelled_at', null),
        // 복원수리 매출 A: 접수시스템 (입금 확인된 건)
        db.from('repairs').select('total_amount, qty_mamoru, qty_other')
          .not('paid_at', 'is', null)
          .gte('created_at', monthISO)
          .not('status', 'eq', 'cancelled'),
        // 복원수리 매출 B: 판매시스템 (이번달, category=RS, 취소 제외) + 고객유형으로 B2B 구분
        db.from('offline_sale_items').select('total_price, quantity, product_name, category, offline_sales!inner(sale_date, cancelled_at, customer_type)')
          .eq('category', 'RS')
          .gte('offline_sales.sale_date', monthStartDate)
          .is('offline_sales.cancelled_at', null),
        // 납품 매출 (이번달, 확정 이상, 취소 제외)
        db.from('deliveries').select('total_amount, discount_amount')
          .gte('delivery_date', monthStartDate)
          .in('status', ['confirmed', 'shipped', 'settled'])
          .is('cancelled_at', null),
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

      const salesRows = (monthSales.data || []) as { total_amount: number; discount_amount: number }[];
      const salesMonthAmount = salesRows.reduce((s, r) => s + ((r.total_amount || 0) - (r.discount_amount || 0)), 0);

      // 납품 매출
      const deliveryRows = (monthDeliveries.data || []) as { total_amount: number; discount_amount: number }[];
      const deliveryMonthAmount = deliveryRows.reduce((s, r) => s + ((r.total_amount || 0) - (r.discount_amount || 0)), 0);

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
          readyToShip: readyToShipRes.count || 0,
          weekRepairTotal: weekRepairRows.length,
          weekRepairMamoru,
          weekRepairOther,
          ...(() => {
            // 접수시스템 매출 — 마모루 1만원, 타사 2만원 고정단가 (배송비 제외)
            const REPAIR_PRICE_MAMORU = 10000;
            const REPAIR_PRICE_OTHER = 20000;
            const repairRows = (monthRepairsPaid.data || []) as Array<{ total_amount: number; qty_mamoru: number; qty_other: number }>;
            let aMamoru = 0, aOther = 0, aMamoruQty = 0, aOtherQty = 0;
            for (const r of repairRows) {
              aMamoru += (r.qty_mamoru || 0) * REPAIR_PRICE_MAMORU;
              aOther += (r.qty_other || 0) * REPAIR_PRICE_OTHER;
              aMamoruQty += r.qty_mamoru || 0;
              aOtherQty += r.qty_other || 0;
            }
            const repairATotal = aMamoru + aOther;

            // 판매시스템 매출 (category=RS) — 마모루/타사/B2B 구분
            const salesRepairRows = (monthSalesRepairItems.data || []) as Array<{ total_price: number; quantity: number; product_name: string; offline_sales: { customer_type: string | null } }>;
            let bMamoru = 0, bOther = 0, bB2B = 0;
            let cMamoru = 0, cOther = 0, cB2B = 0;
            for (const r of salesRepairRows) {
              const ct = r.offline_sales?.customer_type;
              const isB2B = ct === 'dealer' || ct === 'academy';
              const price = r.total_price || 0;
              const qty = r.quantity || 1;
              if (isB2B) { bB2B += price; cB2B += qty; }
              else if ((r.product_name || '').includes('타사')) { bOther += price; cOther += qty; }
              else { bMamoru += price; cMamoru += qty; }
            }
            return {
              monthRepairAmount: repairATotal + bMamoru + bOther + bB2B,
              monthRepairCount: repairRows.length + salesRepairRows.length,
              monthRepairMamoru: { amount: aMamoru + bMamoru, count: aMamoruQty + cMamoru },
              monthRepairOther: { amount: aOther + bOther, count: aOtherQty + cOther },
              monthRepairB2B: { amount: bB2B, count: cB2B },
            };
          })(),
        },
        sales: {
          monthCount: salesRows.length + deliveryRows.length,
          monthAmount: salesMonthAmount + deliveryMonthAmount,
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

      // 이번달 매출 계산용
      const now = new Date();
      const monthStart = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-01`;
      const monthISO = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();

      // 오늘/이번주 작업 일지용
      const todayStart = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}T00:00:00+09:00`;
      const dow = now.getDay();
      const mondayOff = dow === 0 ? -6 : 1 - dow;
      const monday = new Date(now);
      monday.setDate(now.getDate() + mondayOff);
      const weekStartISO = monday.toISOString();
      const weekStartDate = monday.toISOString().slice(0, 10);
      const todayDate = now.toISOString().slice(0, 10);

      const [counts, staleCount, unpaidCount, intakeNewCount, monthRepairsPaid, monthSalesRepairItems, todayCompleted, weekCompleted, weekSalesRepair] = await Promise.all([
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
        // 복원수리 매출 A: 접수시스템 (입금 완료 건)
        supabase
          .from('repairs')
          .select('total_amount, qty_mamoru, qty_other')
          .not('paid_at', 'is', null)
          .gte('created_at', monthISO)
          .not('status', 'eq', 'cancelled'),
        // 복원수리 매출 B: 판매시스템 (이번달, category=RS, 취소 제외)
        (supabase as any)
          .from('offline_sale_items')
          .select('total_price, quantity, product_name, category, offline_sales!inner(sale_date, cancelled_at, customer_type)')
          .eq('category', 'RS')
          .gte('offline_sales.sale_date', monthStart)
          .is('offline_sales.cancelled_at', null),
        // 오늘 작업 완료 (shipped/delivered/completed 상태로 변경된 건)
        (supabase as any)
          .from('repair_history')
          .select('repair_id, to_status, repairs:repair_id(qty_mamoru, qty_other)')
          .in('to_status', ['shipped', 'delivered', 'completed'])
          .gte('created_at', todayStart),
        // 이번주 작업 완료
        (supabase as any)
          .from('repair_history')
          .select('repair_id, to_status, repairs:repair_id(qty_mamoru, qty_other)')
          .in('to_status', ['shipped', 'delivered', 'completed'])
          .gte('created_at', weekStartISO),
        // 이번주 판매시스템 복원수리 (B2B 포함)
        (supabase as any)
          .from('offline_sale_items')
          .select('quantity, product_name, offline_sales!inner(sale_date, cancelled_at, customer_type)')
          .eq('category', 'RS')
          .gte('offline_sales.sale_date', weekStartDate)
          .lte('offline_sales.sale_date', todayDate)
          .is('offline_sales.cancelled_at', null),
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
        // 복원수리 매출 합산 (마모루/타사/B2B 구분)
        ...(() => {
          const REPAIR_PRICE_MAMORU = 10000;
          const REPAIR_PRICE_OTHER = 20000;
          const repairRows = (monthRepairsPaid.data || []) as Array<{ total_amount: number; qty_mamoru: number; qty_other: number }>;
          let aMamoru = 0, aOther = 0, aMamoruQty = 0, aOtherQty = 0;
          for (const r of repairRows) {
            aMamoru += (r.qty_mamoru || 0) * REPAIR_PRICE_MAMORU;
            aOther += (r.qty_other || 0) * REPAIR_PRICE_OTHER;
            aMamoruQty += r.qty_mamoru || 0;
            aOtherQty += r.qty_other || 0;
          }
          const repairATotal = aMamoru + aOther;

          const salesRepairRows = (monthSalesRepairItems.data || []) as Array<{ total_price: number; quantity: number; product_name: string; offline_sales: { customer_type: string | null } }>;
          let bMamoru = 0, bOther = 0, bB2B = 0;
          let cMamoru = 0, cOther = 0, cB2B = 0;
          for (const r of salesRepairRows) {
            const ct = r.offline_sales?.customer_type;
            const isB2B = ct === 'dealer' || ct === 'academy';
            const price = r.total_price || 0;
            const qty = r.quantity || 1;
            if (isB2B) { bB2B += price; cB2B += qty; }
            else if ((r.product_name || '').includes('타사')) { bOther += price; cOther += qty; }
            else { bMamoru += price; cMamoru += qty; }
          }
          return {
            monthRepairAmount: repairATotal + bMamoru + bOther + bB2B,
            monthRepairCount: repairRows.length + salesRepairRows.length,
            monthRepairMamoru: { amount: aMamoru + bMamoru, count: aMamoruQty + cMamoru },
            monthRepairOther: { amount: aOther + bOther, count: aOtherQty + cOther },
            monthRepairB2B: { amount: bB2B, count: cB2B },
          };
        })(),
        // 작업 일지
        todayWork: (() => {
          const rows = (todayCompleted.data || []) as Array<{ repair_id: string; repairs: { qty_mamoru: number; qty_other: number } | null }>;
          const unique = new Set(rows.map(r => r.repair_id));
          let mamoru = 0, other = 0;
          unique.forEach(id => {
            const row = rows.find(r => r.repair_id === id);
            mamoru += row?.repairs?.qty_mamoru || 0;
            other += row?.repairs?.qty_other || 0;
          });
          return { count: unique.size, mamoru, other };
        })(),
        weekWork: (() => {
          // 접수시스템 (출고 이상 상태 전환)
          const rows = (weekCompleted.data || []) as Array<{ repair_id: string; repairs: { qty_mamoru: number; qty_other: number } | null }>;
          const unique = new Set(rows.map(r => r.repair_id));
          let mamoru = 0, other = 0;
          unique.forEach(id => {
            const row = rows.find(r => r.repair_id === id);
            mamoru += row?.repairs?.qty_mamoru || 0;
            other += row?.repairs?.qty_other || 0;
          });
          // 판매시스템 B2B 복원수리
          const salesRows = (weekSalesRepair.data || []) as Array<{ quantity: number; product_name: string; offline_sales: { customer_type: string | null } }>;
          let b2b = 0;
          for (const r of salesRows) {
            const ct = r.offline_sales?.customer_type;
            if (ct === 'dealer' || ct === 'academy') {
              b2b += r.quantity || 1;
            }
          }
          return { count: unique.size, mamoru, other, b2b };
        })(),
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

// ============================================
// 오늘 상담 일정 (confirmed + 오늘 날짜)
// ============================================

export function useTodayConsultations() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['today-consultations'],
    staleTime: 60_000,
    queryFn: async () => {
      const today = new Date().toISOString().slice(0, 10);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any)
        .from('consultations')
        .select('id, name, phone, consultation_type, visit_date, visit_time, status')
        .eq('status', 'confirmed')
        .eq('visit_date', today)
        .order('visit_time', { ascending: true })
        .limit(10);
      if (error) throw error;
      return (data || []) as Array<{
        id: string;
        name: string;
        phone: string | null;
        consultation_type: string;
        visit_date: string;
        visit_time: string | null;
        status: string;
      }>;
    },
  });
}

// ============================================
// 저재고 알림 (stock_quantity 1~3, 재고 사용 상품만)
// ============================================

export function useLowStockAlert() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['low-stock-alert'],
    staleTime: 60_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any)
        .from('products')
        .select('id, name, sku, stock_quantity')
        .gte('stock_quantity', 0)  // 재고 사용 상품만 (-1 제외)
        .lte('stock_quantity', 3)
        .eq('is_active', true)
        .order('stock_quantity', { ascending: true })
        .limit(10);
      if (error) throw error;
      return (data || []) as Array<{
        id: string;
        name: string;
        sku: string;
        stock_quantity: number;
      }>;
    },
  });
}

// ============================================
// 매입 입고대기 알림 (ordered/deposit_paid 상태)
// ============================================

export function usePurchasingAlert() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['purchasing-alert'],
    staleTime: 60_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any)
        .from('purchase_orders')
        .select('id, po_number, supplier_name, total_amount, status, expected_date')
        .in('status', ['ordered', 'deposit_paid'])
        .order('expected_date', { ascending: true })
        .limit(10);
      if (error) throw error;
      return (data || []) as Array<{
        id: string;
        po_number: string;
        supplier_name: string;
        total_amount: number;
        status: string;
        expected_date: string | null;
      }>;
    },
  });
}

// ============================================
// 부자재 주문필요 알림
// ============================================

export function useSuppliesAlert() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['supplies-alert'],
    staleTime: 60_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any)
        .from('supplies')
        .select('id, name, status')
        .eq('status', 'needed')
        .limit(10);
      if (error) throw error;
      return (data || []) as Array<{ id: string; name: string; status: string }>;
    },
  });
}

/** 운송장 잔여 알림 (100건 미만 시 경고) */
export function useWaybillAlert() {
  return useQuery({
    queryKey: ['waybill-alert'],
    staleTime: 300_000, // 5분
    queryFn: async () => {
      const res = await fetch('/api/waybill');
      if (!res.ok) return null;
      const data = await res.json();
      return data as { remaining: number; current_number: number; end_number: number } | null;
    },
  });
}

/** 리뷰 신규 등록 알림 (pending 상태 리뷰) */
export function useNewReviewAlert() {
  const supabase = createClient();
  return useQuery({
    queryKey: ['new-review-alert'],
    staleTime: 60_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, count } = await (supabase as any)
        .from('reviews')
        .select('id, name, stars, created_at', { count: 'exact' })
        .eq('status', 'pending')
        .order('created_at', { ascending: false })
        .limit(5);
      return { reviews: data || [], count: count || 0 };
    },
  });
}
