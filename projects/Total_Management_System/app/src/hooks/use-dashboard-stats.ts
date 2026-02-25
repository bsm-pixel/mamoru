'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Consultation } from '@/lib/supabase/types';

// ============================================
// 허브 대시보드 통계
// ============================================

/** 허브 대시보드: 3개 카테고리별 핵심 수치 */
export function useHubStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['hub-stats'],
    staleTime: 15_000,
    queryFn: async () => {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const todayStr = today.toISOString().slice(0, 10);

      const [payDone, shipping, pendingAdmin, todayVisit, intake, costNotified] =
        await Promise.all([
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'pay_done'),
          supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'shipping'),
          supabase.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'pending_admin'),
          supabase
            .from('consultations')
            .select('*', { count: 'exact', head: true })
            .eq('visit_date', todayStr)
            .in('status', ['confirmed', 'assigned']),
          supabase.from('repairs').select('*', { count: 'exact', head: true }).eq('status', 'intake'),
          supabase.from('repairs').select('*', { count: 'exact', head: true }).eq('status', 'cost_notified'),
        ]);

      return {
        orders: {
          payDone: payDone.count || 0,
          shipping: shipping.count || 0,
        },
        consultations: {
          pendingAdmin: pendingAdmin.count || 0,
          todayVisit: todayVisit.count || 0,
        },
        repairs: {
          intake: intake.count || 0,
          costNotified: costNotified.count || 0,
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
    staleTime: 15_000,
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
    staleTime: 15_000,
    queryFn: async () => {
      const today = new Date();
      const todayStr = today.toISOString().slice(0, 10);

      // 이번주 월~일 범위
      const dayOfWeek = today.getDay();
      const monday = new Date(today);
      monday.setDate(today.getDate() - ((dayOfWeek + 6) % 7));
      const sunday = new Date(monday);
      sunday.setDate(monday.getDate() + 6);
      const monStr = monday.toISOString().slice(0, 10);
      const sunStr = sunday.toISOString().slice(0, 10);

      const [pendingAdmin, todayStore, todayField, suggested, weekVisits, todaySchedule] =
        await Promise.all([
          supabase.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'pending_admin'),
          supabase
            .from('consultations')
            .select('*', { count: 'exact', head: true })
            .eq('visit_date', todayStr)
            .eq('consultation_type', 'store_visit'),
          supabase
            .from('consultations')
            .select('*', { count: 'exact', head: true })
            .eq('visit_date', todayStr)
            .eq('consultation_type', 'field_request'),
          supabase.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'suggested'),
          supabase
            .from('consultations')
            .select('*', { count: 'exact', head: true })
            .gte('visit_date', monStr)
            .lte('visit_date', sunStr)
            .in('status', ['confirmed', 'assigned', 'in_progress']),
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
        pendingAdmin: pendingAdmin.count || 0,
        todayStore: todayStore.count || 0,
        todayField: todayField.count || 0,
        suggested: suggested.count || 0,
        weekVisits: weekVisits.count || 0,
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
    staleTime: 15_000,
    queryFn: async () => {
      // 3일 전 기준
      const threeDaysAgo = new Date();
      threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
      const threeDaysAgoISO = threeDaysAgo.toISOString();

      const statuses = [
        'intake', 'pickup_scheduled', 'inspecting',
        'cost_notified', 'payment_confirmed', 'repairing',
        'ready_to_ship', 'shipped',
      ] as const;

      const [counts, staleCount] = await Promise.all([
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
        // 경과일 3일 이상 미처리
        supabase
          .from('repairs')
          .select('*', { count: 'exact', head: true })
          .in('status', ['inspecting', 'cost_notified'])
          .lt('updated_at', threeDaysAgoISO),
      ]);

      const byStatus = Object.fromEntries(counts.map((c) => [c.status, c.count])) as Record<string, number>;

      return {
        byStatus,
        staleCount: staleCount.count || 0,
        pipeline: [
          { label: '접수', count: byStatus.intake || 0, status: 'intake' },
          { label: '수거대기', count: byStatus.pickup_scheduled || 0, status: 'pickup_scheduled' },
          { label: '검수', count: byStatus.inspecting || 0, status: 'inspecting' },
          { label: '비용안내', count: byStatus.cost_notified || 0, status: 'cost_notified' },
          { label: '입금확인', count: byStatus.payment_confirmed || 0, status: 'payment_confirmed' },
          { label: '수리중', count: byStatus.repairing || 0, status: 'repairing' },
          { label: '출고대기', count: byStatus.ready_to_ship || 0, status: 'ready_to_ship' },
          { label: '배송중', count: byStatus.shipped || 0, status: 'shipped' },
        ],
      };
    },
  });
}
