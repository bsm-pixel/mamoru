'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Repair } from '@/lib/supabase/types';
import type { RepairTabKey, RepairTabDef } from '@/components/repairs/repair-tab-bar';

interface RepairTabData {
  intake: Repair[];
  pickup_needed: Repair[];
  inbound_waiting: Repair[];
  in_progress: Repair[];
  ready_to_ship: Repair[];
  shipped: Repair[];
}

/** 복원수리 대시보드 탭별 데이터 + 카운트 */
export function useRepairTabData() {
  const supabase = createClient();

  const query = useQuery({
    queryKey: ['repair-tabs'],
    staleTime: 10_000,
    queryFn: async () => {
      const threeDaysAgo = new Date();
      threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
      const threeDaysAgoISO = threeDaysAgo.toISOString();

      // 병렬 쿼리 6개 + staleCount
      const [
        intakeRes,
        pickupNeededRes,
        inboundWaitingDirectRes,
        inboundWaitingPickupRes,
        inProgressRes,
        readyToShipRes,
        shippedRes,
        staleRes,
      ] = await Promise.all([
        // 1) 신규접수: intake + confirmed_at IS NULL
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'intake')
          .is('confirmed_at', null)
          .order('received_at', { ascending: false })
          .limit(50),

        // 2) 수거접수필요: intake + 방문수거 + confirmed_at IS NULL
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'intake')
          .eq('proceed_type', '방문수거')
          .is('confirmed_at', null)
          .order('received_at', { ascending: false })
          .limit(50),

        // 3a) 입고대기 — 직접발송 확인완료: intake + 방문수거 아닌 것 + confirmed_at IS NOT NULL
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'intake')
          .not('proceed_type', 'eq', '방문수거')
          .not('confirmed_at', 'is', null)
          .order('received_at', { ascending: false })
          .limit(50),

        // 3b) 입고대기 — 방문수거 수거완료: pickup_scheduled
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'pickup_scheduled')
          .order('received_at', { ascending: false })
          .limit(50),

        // 4) 진행중: cost_notified + repairing
        supabase
          .from('repairs')
          .select('*')
          .in('status', ['cost_notified', 'repairing'])
          .order('received_at', { ascending: false })
          .limit(50),

        // 5) 출고대기: ready_to_ship
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'ready_to_ship')
          .order('received_at', { ascending: false })
          .limit(50),

        // 6) 출고완료: shipped + delivered + completed
        supabase
          .from('repairs')
          .select('*')
          .in('status', ['shipped', 'delivered', 'completed'])
          .order('shipped_at', { ascending: false })
          .limit(50),

        // 경과일 3일 이상 미처리
        supabase
          .from('repairs')
          .select('*', { count: 'exact', head: true })
          .in('status', ['intake', 'cost_notified'])
          .lt('updated_at', threeDaysAgoISO),
      ]);

      const intake = (intakeRes.data || []) as Repair[];
      const pickupNeeded = (pickupNeededRes.data || []) as Repair[];
      const inboundDirect = (inboundWaitingDirectRes.data || []) as Repair[];
      const inboundPickup = (inboundWaitingPickupRes.data || []) as Repair[];
      const inboundWaiting = [...inboundDirect, ...inboundPickup].sort(
        (a, b) => new Date(b.received_at).getTime() - new Date(a.received_at).getTime()
      );
      const inProgress = (inProgressRes.data || []) as Repair[];
      const readyToShip = (readyToShipRes.data || []) as Repair[];
      const shipped = (shippedRes.data || []) as Repair[];

      return {
        tabData: {
          intake,
          pickup_needed: pickupNeeded,
          inbound_waiting: inboundWaiting,
          in_progress: inProgress,
          ready_to_ship: readyToShip,
          shipped,
        } as RepairTabData,
        staleCount: staleRes.count || 0,
      };
    },
  });

  const tabData = query.data?.tabData || {
    intake: [],
    pickup_needed: [],
    inbound_waiting: [],
    in_progress: [],
    ready_to_ship: [],
    shipped: [],
  };

  const tabs: RepairTabDef[] = [
    { key: 'intake', label: '신규접수', count: tabData.intake.length },
    { key: 'pickup_needed', label: '수거접수필요', count: tabData.pickup_needed.length },
    { key: 'inbound_waiting', label: '입고대기', count: tabData.inbound_waiting.length },
    { key: 'in_progress', label: '진행중', count: tabData.in_progress.length },
    { key: 'ready_to_ship', label: '출고대기', count: tabData.ready_to_ship.length },
    { key: 'shipped', label: '출고완료', count: tabData.shipped.length },
  ];

  return {
    tabs,
    tabData,
    isLoading: query.isLoading,
    staleCount: query.data?.staleCount || 0,
  };
}
