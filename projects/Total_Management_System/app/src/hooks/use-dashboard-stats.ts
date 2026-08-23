'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import { toLocalDateString } from '@/lib/utils/format';
import { deliveryNet } from '@/lib/sales/amounts';
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
    monthRepairAmount: number;   // 복원수리 매출 전체 (A 접수시스템 + B 판매RS + C 납품RS)
    monthRepairCount: number;    // 복원수리 수량 전체 (자루 기준)
    monthRepairAOnly: number;    // 접수시스템(A채널) 매출만 — 월 목표 계산용 (B/C채널 RS는 sales.monthAmount에 이미 포함되므로 중복 방지)
    monthRepairMamoru: { amount: number; count: number };
    monthRepairOther: { amount: number; count: number };
    monthRepairB2B: { amount: number; count: number };
  };
  sales: {
    monthCount: number;
    monthAmount: number;       // 오프라인판매(RS 포함) + 납품(RS 포함) 전체 — 호환 유지 (오프라인 판매 카드용)
    salesB2C: number;          // B2C 제품 매출 = offline_sales(소매) total−discount − 그 주문 RS_total
    salesB2B: number;          // B2B 제품 매출 = offline_sales(딜러/아카데미) total−discount − RS_total + 납품 total−discount − 납품 RS_total
    salesOnline: number;       // 126: 아임웹 온라인 주문 제품매출 (B2C에 합산됨. 리포트 summary와 동일 기준)
  };
}

/** 이번달 아임웹 온라인 주문 제품매출 (order_items total, 배송비 제외, 결제완료 이상 = 취소/환불/무통장미입금 제외).
 *  리포트(/api/reports/summary)와 동일 기준. 대시보드에서 B2C 제품매출에 합산된다. */
async function getMonthOnlineRevenue(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  db: any, monthStartISO: string,
): Promise<number> {
  const { data: ords } = await db.from('orders').select('id')
    .gte('ordered_at', monthStartISO)
    .not('status', 'in', '("cancelled","refunded","pay_wait")');
  const ids = (ords || []).map((o: { id: string }) => o.id);
  if (ids.length === 0) return 0;
  const { data: items } = await db.from('order_items')
    .select('product_name, total_price, unit_price, quantity').in('order_id', ids);
  return (items || [])
    .filter((it: { product_name: string }) => it.product_name !== '배송비')
    .reduce((s: number, it: { total_price: number; unit_price: number; quantity: number }) =>
      s + (it.total_price || (it.unit_price || 0) * (it.quantity || 0)), 0);
}

/** R3+P12: 허브 대시보드 — RPC 1회 호출로 통합 (fallback: 기존 14개 쿼리) */
export function useHubStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['hub-stats'],
    staleTime: 30_000, // 30초 — RPC 1회면 충분
    queryFn: async (): Promise<HubStatsResult> => {
      // 이번달 아임웹 온라인 주문 제품매출 — RPC/fallback 어느 경로든 B2C 매출에 합산 (2026-08-20)
      const monthStartISO = new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString();
      const monthOnline = await getMonthOnlineRevenue(supabase, monthStartISO);

      // RPC 호출 시도 (018_hub_stats_rpc.sql 배포 후 동작)
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data: rpcData, error: rpcError } = await (supabase as any).rpc('get_hub_stats');

      if (!rpcError && rpcData) {
        const d = rpcData as HubStatsResult;

        // RPC 077/078 — get_hub_stats 가 복원수리 매출(A 접수 + B 판매RS + C 납품RS) / 카운트(수량) /
        //   monthRepairAOnly / monthRepairB2B(offline B2B + 납품RS) 를 모두 계산한다.
        //   RPC 078+ 는 sales.salesB2C/salesB2B (제품 매출, RS 제외) + deliveries 포함 monthCount/monthAmount 도 반환.
        //   RPC 077(078 미배포) 일 때만 후처리에서 deliveries 추가 + 제품매출 B2C/B2B 분리를 직접 쿼리한다.
        //   ⚠️ 복원수리(monthRepairAmount 등) 에 납품 RS(dlRs) 를 또 더하면 C채널 중복 계상 — 절대 더하지 말 것.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const dlDb = supabase as any;
        // 월 시작 KST 명시 변환 (toISOString은 UTC라 KST 5/1 자정→4/30으로 잘못 변환되는 버그 회피)
        const msd = toLocalDateString(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
        const ds = d.sales as { monthCount?: number; monthAmount?: number; salesB2C?: number; salesB2B?: number };
        let salesMonthCount: number, salesMonthAmount: number, salesB2C: number, salesB2B: number;

        if (ds.salesB2C !== undefined && ds.salesB2B !== undefined) {
          // RPC 078+ : sales 객체가 제품매출 분리 + deliveries 포함 값을 모두 반환 → 그대로 사용
          salesMonthCount = ds.monthCount ?? 0;
          salesMonthAmount = ds.monthAmount ?? 0;
          salesB2C = ds.salesB2C;
          salesB2B = ds.salesB2B;
        } else {
          // RPC 077 : 후처리에서 deliveries 추가 + 제품매출 B2C/B2B 분리 (RS 제외) 직접 계산
          const isB2BCt = (ct: string | null | undefined) => ct === 'dealer' || ct === 'academy';
          const [dlRes, osRes, osRsRes, dlRsRes] = await Promise.all([
            dlDb.from('deliveries').select('total_amount, discount_amount')
              .gte('delivery_date', msd).in('status', ['confirmed', 'shipped', 'settled']).is('cancelled_at', null),
            dlDb.from('offline_sales').select('total_amount, discount_amount, customer_type')
              .gte('sale_date', msd).is('cancelled_at', null).is('returned_at', null),
            dlDb.from('offline_sale_items').select('total_price, offline_sales!inner(sale_date, cancelled_at, customer_type)')
              .eq('category', 'RS').gt('total_price', 0).gte('offline_sales.sale_date', msd).is('offline_sales.cancelled_at', null).is('offline_sales.returned_at', null),
            (async () => {
              const { data: dlIds } = await dlDb.from('deliveries').select('id')
                .gte('delivery_date', msd).in('status', ['confirmed', 'shipped', 'settled']).is('cancelled_at', null);
              if (!dlIds || dlIds.length === 0) return { data: [] };
              return dlDb.from('delivery_items').select('total_price').eq('category', 'RS').gt('total_price', 0)
                .in('delivery_id', dlIds.map((x: { id: string }) => x.id));
            })(),
          ]);
          const dlRows = (dlRes.data || []) as { total_amount: number; discount_amount?: number }[];
          const dlAmount = dlRows.reduce((s, r) => s + deliveryNet(r), 0); // 납품 total은 이미 net(할인 재차감 금지)
          let osB2C = 0, osB2B = 0;
          for (const r of (osRes.data || []) as { total_amount: number; discount_amount?: number; customer_type: string | null }[]) {
            const net = (r.total_amount || 0) - (r.discount_amount || 0);
            if (isB2BCt(r.customer_type)) osB2B += net; else osB2C += net;
          }
          let osRsB2C = 0, osRsB2B = 0;
          for (const r of (osRsRes.data || []) as { total_price: number; offline_sales: { customer_type: string | null } }[]) {
            if (isB2BCt(r.offline_sales?.customer_type)) osRsB2B += (r.total_price || 0); else osRsB2C += (r.total_price || 0);
          }
          const dlRs = ((dlRsRes.data || []) as { total_price: number }[]).reduce((s, r) => s + (r.total_price || 0), 0);
          salesMonthCount = (ds.monthCount ?? 0) + dlRows.length;
          salesMonthAmount = (ds.monthAmount ?? 0) + dlAmount;
          salesB2C = osB2C - osRsB2C;
          salesB2B = (osB2B - osRsB2B) + (dlAmount - dlRs);
        }

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
            monthRepairAmount: d.repairs?.monthRepairAmount ?? 0,  // RPC 077 = A(접수)+B(판매RS)+C(납품RS) 전체
            monthRepairCount: d.repairs?.monthRepairCount ?? 0,    // RPC 077 = A+B+C 수량(자루) 기준
            monthRepairAOnly: d.repairs?.monthRepairAOnly ?? 0,    // RPC 077 = 접수시스템 A채널 매출만 (077 미배포 시 0)
            monthRepairMamoru: d.repairs?.monthRepairMamoru ?? { amount: 0, count: 0 },
            monthRepairOther: d.repairs?.monthRepairOther ?? { amount: 0, count: 0 },
            monthRepairB2B: d.repairs?.monthRepairB2B ?? { amount: 0, count: 0 },  // RPC 077 = offline B2B + 납품RS 전체
          },
          sales: {
            monthCount: salesMonthCount,
            monthAmount: salesMonthAmount,
            salesB2C,
            salesB2B,
            salesOnline: monthOnline,
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
      const monthStartDate = toLocalDateString(monthStart);
      const [
        payDone, preparing, shipping, delivered,
        weekOrders, monthOrders,
        newConsult, confirmedConsult, needAction,
        intakeNew, repairPending, repairWorking, readyToShipRes,
        weekRepairs,
        monthSales,
        monthRepairsPaid, monthSalesRepairItems,
        monthDeliveries,
        monthDeliveryRepairItems,
      ] = await Promise.all([
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'pay_done'),
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'preparing'),
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'shipping'),
        db.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'delivered'),
        db.from('orders').select('paid_amount').gte('ordered_at', monISO).not('status', 'in', '("cancelled","refunded")'),
        db.from('orders').select('paid_amount').gte('ordered_at', monthISO).not('status', 'in', '("cancelled","refunded")'),
        db.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'pending_admin'),
        db.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'confirmed'),
        // 075: pending_admin 제거 — 신규 상담이 재요청으로 중복 카운트되던 버그 fix
        db.from('consultations').select('*', { count: 'exact', head: true })
          .in('status', ['reschedule_requested', 'change_requested'])
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
        db.from('offline_sales').select('total_amount, discount_amount, customer_type')
          .gte('sale_date', monthStartDate)
          .is('cancelled_at', null)
          .is('returned_at', null),
        // 075: 복원수리 매출 A — 옵션 A "발생 기준" 적용 (paid_at 조건 제거 → 미입금도 매출 발생으로 카운트)
        db.from('repairs').select('total_amount, qty_mamoru, qty_other')
          .gte('created_at', monthISO)
          .not('status', 'eq', 'cancelled'),
        // 075: 복원수리 매출 B — 판매시스템 (category=RS, 취소 제외, 0원 무상 제외) + 고객유형으로 B2B 구분
        db.from('offline_sale_items').select('total_price, quantity, product_name, category, offline_sales!inner(sale_date, cancelled_at, customer_type)')
          .eq('category', 'RS')
          .gt('total_price', 0)
          .gte('offline_sales.sale_date', monthStartDate)
          .is('offline_sales.cancelled_at', null)
          .is('offline_sales.returned_at', null),
        // 납품 매출 + 항목 (이번달, 확정 이상, 취소 제외)
        db.from('deliveries').select('total_amount, discount_amount')
          .gte('delivery_date', monthStartDate)
          .in('status', ['confirmed', 'shipped', 'settled'])
          .is('cancelled_at', null),
        // 납품 복원수리 수량 — 간단 쿼리 (납품 ID로 필터)
        (async () => {
          const { data: dlIds } = await db.from('deliveries').select('id')
            .gte('delivery_date', monthStartDate)
            .in('status', ['confirmed', 'shipped', 'settled'])
            .is('cancelled_at', null);
          if (!dlIds || dlIds.length === 0) return { data: [] };
          const ids = dlIds.map((d: { id: string }) => d.id);
          return db.from('delivery_items').select('quantity, total_price, product_name').eq('category', 'RS').gt('total_price', 0).in('delivery_id', ids);
        })(),
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

      const salesRows = (monthSales.data || []) as { total_amount: number; discount_amount: number; customer_type?: string | null }[];
      const salesMonthAmount = salesRows.reduce((s, r) => s + ((r.total_amount || 0) - (r.discount_amount || 0)), 0);

      // 납품 매출
      const deliveryRows = (monthDeliveries.data || []) as { total_amount: number; discount_amount: number }[];
      const deliveryMonthAmount = deliveryRows.reduce((s, r) => s + deliveryNet(r), 0); // 납품 total은 이미 net

      // 납품 복원수리 B2B 수량 (배송비 항목은 자루 수에서 제외 — 금액엔 포함)
      const deliveryRepairRows = (monthDeliveryRepairItems.data || []) as { quantity: number; total_price: number; product_name?: string }[];
      const deliveryRepairB2BQty = deliveryRepairRows.reduce((s, r) => s + ((r.product_name || '') === '배송비' ? 0 : (r.quantity || 0)), 0);
      const deliveryRepairAmount = deliveryRepairRows.reduce((s, r) => s + (r.total_price || 0), 0);

      // 2단계 매출 3분할 — 제품 매출 B2C/B2B 분리 (RS 제외)
      //   B2C 제품 = offline_sales(소매/온라인) (total−discount) − 그 채널 RS_total
      //   B2B 제품 = offline_sales(딜러/아카데미) (total−discount) − RS_total + 납품 전체 (total−discount − 납품 RS_total)
      const isB2BCt = (ct: string | null | undefined) => ct === 'dealer' || ct === 'academy';
      const osProductB2C = salesRows.reduce((s, r) => s + (isB2BCt(r.customer_type) ? 0 : (r.total_amount || 0) - (r.discount_amount || 0)), 0);
      const osProductB2B = salesRows.reduce((s, r) => s + (isB2BCt(r.customer_type) ? (r.total_amount || 0) - (r.discount_amount || 0) : 0), 0);
      const osRepairRowsForSplit = (monthSalesRepairItems.data || []) as Array<{ total_price: number; offline_sales: { customer_type: string | null } }>;
      const osRsB2C = osRepairRowsForSplit.reduce((s, r) => s + (isB2BCt(r.offline_sales?.customer_type) ? 0 : (r.total_price || 0)), 0);
      const osRsB2B = osRepairRowsForSplit.reduce((s, r) => s + (isB2BCt(r.offline_sales?.customer_type) ? (r.total_price || 0) : 0), 0);
      const salesB2C = osProductB2C - osRsB2C;
      const salesB2B = (osProductB2B - osRsB2B) + (deliveryMonthAmount - deliveryRepairAmount);

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
            // 접수시스템 매출 — 실제청구액(total_amount) 을 단가 비중으로 마모루/타사 안분 (2026-05-19)
            //   배송비·가공비 포함 + 분리합 = total_amount 일치 (회계 금액 기준과 통일)
            const REPAIR_PRICE_MAMORU = 10000;
            const REPAIR_PRICE_OTHER = 20000;
            const repairRows = (monthRepairsPaid.data || []) as Array<{ total_amount: number; qty_mamoru: number; qty_other: number }>;
            let aMamoru = 0, aOther = 0, aMamoruQty = 0, aOtherQty = 0;
            for (const r of repairRows) {
              const qm = r.qty_mamoru || 0;
              const qo = r.qty_other || 0;
              const total = r.total_amount || 0;
              const baseM = qm * REPAIR_PRICE_MAMORU;
              const baseO = qo * REPAIR_PRICE_OTHER;
              const baseSum = baseM + baseO;
              if (baseSum > 0) {
                const mPart = Math.round(total * (baseM / baseSum));
                aMamoru += mPart;
                aOther += total - mPart;
              } else if (total > 0) {
                aMamoru += total;
              }
              aMamoruQty += qm;
              aOtherQty += qo;
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
              // 배송비 항목은 매출(amount)엔 포함하되 자루 수(qty)엔 포함하지 않음
              const qty = (r.product_name || '') === '배송비' ? 0 : (r.quantity || 1);
              if (isB2B) { bB2B += price; cB2B += qty; }
              else if ((r.product_name || '').includes('타사')) { bOther += price; cOther += qty; }
              else { bMamoru += price; cMamoru += qty; }
            }
            return {
              // 복원수리 매출 전체 = A(접수시스템) + B(판매RS) + C(납품RS)
              monthRepairAmount: repairATotal + bMamoru + bOther + bB2B + deliveryRepairAmount,
              // 복원수리 수량 전체 (자루) = A(repairs qty) + B(offline RS qty) + C(deliveries RS qty)
              monthRepairCount: aMamoruQty + aOtherQty + cMamoru + cOther + cB2B + deliveryRepairB2BQty,
              // 접수시스템 A채널 매출만 — sales.monthAmount에 없는 유일한 복원수리 매출 (월 목표 계산용)
              monthRepairAOnly: repairATotal,
              monthRepairMamoru: { amount: aMamoru + bMamoru, count: aMamoruQty + cMamoru },
              monthRepairOther: { amount: aOther + bOther, count: aOtherQty + cOther },
              monthRepairB2B: { amount: bB2B + deliveryRepairAmount, count: cB2B + deliveryRepairB2BQty },
            };
          })(),
        },
        sales: {
          monthCount: salesRows.length + deliveryRows.length,
          monthAmount: salesMonthAmount + deliveryMonthAmount,
          salesB2C,
          salesB2B,
          salesOnline: monthOnline,
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

      const [payDone, preparing, readyToShip, shipping, delivered, todayOrders] =
        await Promise.all([
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'pay_done'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'preparing'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'ready_to_ship'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'shipping'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'delivered'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).gte('ordered_at', todayISO),
        ]);

      return {
        payDone: payDone.count || 0,
        preparing: preparing.count || 0,
        readyToShip: readyToShip.count || 0,
        shipping: shipping.count || 0,
        delivered: delivered.count || 0,
        todayOrders: todayOrders.count || 0,
        pipeline: [
          { label: '결제완료', count: payDone.count || 0, status: 'pay_done' },
          { label: '준비중', count: preparing.count || 0, status: 'preparing' },
          { label: '배송대기', count: readyToShip.count || 0, status: 'ready_to_ship' },
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
      const todayStr = toLocalDateString(now);
      // 1달 전
      const oneMonthAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000).toISOString();

      // 상태 기반 4버킷 — 각 카드가 겹치지 않도록 상태로 명확히 분리 (2026-07-01, 6시간 휴리스틱 폐기)
      const [
        newIntake,
        inCoordination,
        confirmed,
        completedMonth,
        todaySchedule,
      ] = await Promise.all([
        // 신규 = 미확인 접수 (관리자 최초 대응 필요)
        supabase
          .from('consultations')
          .select('*', { count: 'exact', head: true })
          .eq('status', 'pending_admin'),
        // 조율중 = 활성 상태 중 신규(pending_admin)·확정(confirmed)·완료·취소 제외
        //   → suggested / assigned / reschedule_requested / change_requested / in_progress / on_hold
        supabase
          .from('consultations')
          .select('*', { count: 'exact', head: true })
          .not('status', 'in', '("pending_admin","confirmed","completed","cancelled")'),
        // 확정 = 잡힌 일정 (조율 끝, 방문 예정)
        supabase
          .from('consultations')
          .select('*', { count: 'exact', head: true })
          .eq('status', 'confirmed'),
        // 완료 (최근 1달)
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
        inCoordination: inCoordination.count || 0,
        confirmed: confirmed.count || 0,
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
      const weekStartDate = toLocalDateString(monday);
      const todayDate = toLocalDateString(now);

      const [counts, staleCount, unpaidCount, intakeNewCount, monthRepairsPaid, monthSalesRepairItems, todayCompleted, weekCompleted, weekSalesRepair, dlRepairItems] = await Promise.all([
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
        // 미입금 건수 (cost_notified 이후 + paid_at IS NULL + 받을 금액>0)
        //   delivered/completed 포함 — 자동배송완료로 완료된 미입금 건이 누락되던 사각지대 fix (2026-06-11)
        supabase
          .from('repairs')
          .select('*', { count: 'exact', head: true })
          .in('status', ['cost_notified', 'repairing', 'ready_to_ship', 'shipped', 'delivered', 'completed'])
          .gt('total_amount', 0)
          .is('paid_at', null),
        // R1: 신규접수 (intake + confirmed_at IS NULL)
        supabase
          .from('repairs')
          .select('*', { count: 'exact', head: true })
          .eq('status', 'intake')
          .is('confirmed_at', null),
        // 복원수리 매출 A: 접수시스템 — 옵션A(발생 기준, 미입금 포함). useHubStats와 동일 정책으로 통일
        //   (거래처/고객이 미입금으로 몰아서 입금하는 케이스가 있어 paid_at 조건 제거)
        supabase
          .from('repairs')
          .select('total_amount, qty_mamoru, qty_other')
          .gte('created_at', monthISO)
          .not('status', 'eq', 'cancelled'),
        // 복원수리 매출 B: 판매시스템 (이번달, category=RS, 취소 제외)
        (supabase as any)
          .from('offline_sale_items')
          .select('total_price, quantity, product_name, category, offline_sales!inner(sale_date, cancelled_at, customer_type)')
          .eq('category', 'RS')
          .gte('offline_sales.sale_date', monthStart)
          .is('offline_sales.cancelled_at', null)
          .is('offline_sales.returned_at', null),
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
          .is('offline_sales.cancelled_at', null)
          .is('offline_sales.returned_at', null),
        // 납품 복원수리 B2B 수량 (delivery_items category=RS)
        (async () => {
          const { data: dlIds } = await (supabase as any).from('deliveries').select('id')
            .gte('delivery_date', monthStart).in('status', ['confirmed', 'shipped', 'settled']).is('cancelled_at', null);
          if (!dlIds || dlIds.length === 0) return { data: [] };
          return (supabase as any).from('delivery_items').select('quantity, total_price, product_name').eq('category', 'RS').gt('total_price', 0).in('delivery_id', dlIds.map((x: { id: string }) => x.id));
        })(),
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
            const qm = r.qty_mamoru || 0;
            const qo = r.qty_other || 0;
            const total = r.total_amount || 0;
            // 실제청구액(total_amount = 수리비+배송비+가공비)을 단가 비중으로 마모루/타사 안분 (2026-05-19)
            // → 배송비·날변형 가공비 포함 + 분리합 = total_amount 정확 일치 (회계 금액 기준과 통일)
            const baseM = qm * REPAIR_PRICE_MAMORU;
            const baseO = qo * REPAIR_PRICE_OTHER;
            const baseSum = baseM + baseO;
            if (baseSum > 0) {
              const mPart = Math.round(total * (baseM / baseSum));
              aMamoru += mPart;
              aOther += total - mPart; // 잔액 = 타사 (반올림 오차 흡수 → 분리합 정확)
            } else if (total > 0) {
              aMamoru += total; // qty 둘 다 0인 이상치 → 전액 마모루로 흡수
            }
            aMamoruQty += qm;
            aOtherQty += qo;
          }
          const repairATotal = aMamoru + aOther;

          const salesRepairRows = (monthSalesRepairItems.data || []) as Array<{ total_price: number; quantity: number; product_name: string; offline_sales: { customer_type: string | null } }>;
          let bMamoru = 0, bOther = 0, bB2B = 0;
          let cMamoru = 0, cOther = 0, cB2B = 0;
          for (const r of salesRepairRows) {
            const ct = r.offline_sales?.customer_type;
            const isB2B = ct === 'dealer' || ct === 'academy';
            const price = r.total_price || 0;
            // 배송비 항목은 매출엔 포함, 자루 수엔 제외
            const qty = (r.product_name || '') === '배송비' ? 0 : (r.quantity || 1);
            if (isB2B) { bB2B += price; cB2B += qty; }
            else if ((r.product_name || '').includes('타사')) { bOther += price; cOther += qty; }
            else { bMamoru += price; cMamoru += qty; }
          }
          const dlRepairAmount = ((dlRepairItems.data || []) as { total_price: number }[]).reduce((s, r) => s + (r.total_price || 0), 0);
          const dlRepairQty = ((dlRepairItems.data || []) as { quantity: number; product_name?: string }[]).reduce((s, r) => s + ((r.product_name || '') === '배송비' ? 0 : (r.quantity || 0)), 0);
          return {
            monthRepairAmount: repairATotal + bMamoru + bOther + bB2B + dlRepairAmount,
            // 복원수리 수량 전체 (자루) = A(repairs qty) + B(offline RS qty) + C(deliveries RS qty). useHubStats와 동일 기준
            monthRepairCount: aMamoruQty + aOtherQty + cMamoru + cOther + cB2B + dlRepairQty,
            monthRepairMamoru: { amount: aMamoru + bMamoru, count: aMamoruQty + cMamoru },
            monthRepairOther: { amount: aOther + bOther, count: aOtherQty + cOther },
            monthRepairB2B: {
              amount: bB2B + dlRepairAmount,
              count: cB2B + dlRepairQty,
            },
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
            if ((ct === 'dealer' || ct === 'academy') && (r.product_name || '') !== '배송비') {
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
      const today = toLocalDateString(new Date());
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
// 저재고 알림 (설정값 기준, 재고 사용 상품만)
// ============================================

export function useLowStockAlert() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['low-stock-alert'],
    staleTime: 60_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;

      // 설정에서 저재고 기준수량 읽기
      const { data: setting } = await db
        .from('system_settings')
        .select('value')
        .eq('key', 'inventory.low_stock_threshold')
        .single();
      const threshold = setting?.value ? (typeof setting.value === 'number' ? setting.value : parseInt(String(setting.value)) || 3) : 3;

      const { data, error } = await db
        .from('products')
        .select('id, name, sku, stock_quantity')
        .gte('stock_quantity', 0)  // 재고 사용 상품만 (-1 제외)
        .lte('stock_quantity', threshold)
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
