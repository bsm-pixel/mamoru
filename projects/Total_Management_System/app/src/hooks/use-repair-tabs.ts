'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Repair } from '@/lib/supabase/types';
import type { RepairTabDef } from '@/components/repairs/repair-tab-bar';

interface RepairTabData {
  intake: Repair[];
  pickup_needed: Repair[];
  visit_scheduled: Repair[];
  inbound_waiting: Repair[];
  in_progress: Repair[];
  ready_to_ship: Repair[];
  shipped: Repair[];
  recall: Repair[];
}

// ── 출고완료 탭 (2026-11-10) ──
export const SHIPPED_WINDOW_DAYS = 30;
const SHIPPED_STATUSES = ['shipped', 'delivered', 'completed'];
/** 끝난 시각 — 택배=배송완료·출고, 방문=전달(delivered_at). 둘 다 없으면 마지막 수정 */
export function repairEndedAt(r: Repair): number {
  return new Date(r.delivered_at || r.shipped_at || r.updated_at).getTime();
}
/** 재수거 접수 후 재작업 전 = '재수리' 탭 소속 (출고완료에서 제외) */
const isRecallPending = (r: Repair) => !!r.recall_booked_at && !r.reworked_at;
const sortByEndedAt = (list: Repair[]) => [...list].sort((a, b) => repairEndedAt(b) - repairEndedAt(a));

/**
 * 출고완료 탭 전체 기간 검색 — 탭은 최근 30일만 불러오므로 이전 고객은 서버에서 찾는다 (이름·전화·접수번호)
 * 2글자 이상일 때만. PostgREST or-문법을 깨는 문자(, ( ) * %)는 제거
 */
export function useRepairShippedSearch(term: string, enabled: boolean) {
  const supabase = createClient();
  const q = term.trim().replace(/[,()*%\\]/g, '');
  const digits = q.replace(/\D/g, '');
  return useQuery({
    queryKey: ['repair-tabs', 'shipped-search', q],
    enabled: enabled && q.length >= 2,
    staleTime: 20_000,
    queryFn: async () => {
      const ors = [`name.ilike.*${q}*`, `as_id.ilike.*${q}*`];
      if (digits.length >= 3) ors.push(`phone_normalized.ilike.*${digits}*`, `phone.ilike.*${digits}*`);
      const { data, error } = await supabase
        .from('repairs')
        .select('*')
        .in('status', SHIPPED_STATUSES)
        .or(ors.join(','))
        .limit(200);
      if (error) throw error;
      return sortByEndedAt(((data || []) as Repair[]).filter((r) => !isRecallPending(r)));
    },
  });
}

/** 복원수리 대시보드 탭별 데이터 + 카운트 */
export function useRepairTabData() {
  const supabase = createClient();

  const query = useQuery({
    queryKey: ['repair-tabs'],
    staleTime: 20_000, // 20초 — 10초는 너무 공격적
    // 170: 우체국 자동접수 진행 중(3분 잠금)인 건이 있을 때만 5초마다 다시 읽음 → 끝나면 목록이 바로 입고대기로
    refetchInterval: (q) => {
      const rows = (q.state.data as { tabData?: RepairTabData } | undefined)?.tabData?.pickup_needed || [];
      const inFlight = rows.some((r) => !r.epost_reservation_no && r.epost_requested_at && Date.now() - new Date(r.epost_requested_at).getTime() < 3 * 60 * 1000);
      return inFlight ? 5000 : false;
    },
    queryFn: async () => {
      const threeDaysAgo = new Date();
      threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
      const threeDaysAgoISO = threeDaysAgo.toISOString();
      const shippedSince = new Date();
      shippedSince.setDate(shippedSince.getDate() - SHIPPED_WINDOW_DAYS);
      const shippedSinceISO = shippedSince.toISOString();

      // 병렬 쿼리 6개 + staleCount
      const [
        intakeRes,
        pickupNeededRes,
        visitScheduledRes,
        inboundWaitingDirectRes,
        inboundWaitingPickupRes,
        inProgressRes,
        readyToShipRes,
        shippedRes,
        recallRes,
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

        // 2) 수거접수필요: intake + 방문수거 + confirmed_at IS NOT NULL
        //    = 접수확인 완료 후 수거 예약이 필요한 방문수거 건 (2026-05-19 fix:
        //      이전엔 confirmed_at IS NULL 이라 접수확인 누르면 orphan 으로 사라졌음)
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'intake')
          .eq('proceed_type', '방문수거')
          .not('confirmed_at', 'is', null)
          .order('received_at', { ascending: false })
          .limit(50),

        // 2b) 방문예정: intake + 직접방문 + confirmed_at IS NOT NULL
        //     = 접수확인 완료 후 고객 매장방문 대기 (2026-07-28: 이전엔 입고대기 3a로 뭉쳐 있었음)
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'intake')
          .eq('proceed_type', '직접방문')
          .not('confirmed_at', 'is', null)
          .order('visit_date', { ascending: true, nullsFirst: false })
          .limit(50),

        // 3a) 입고대기 — 직접발송 확인완료: intake + 방문수거·직접방문 아닌 것 + confirmed_at IS NOT NULL
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'intake')
          .not('proceed_type', 'eq', '방문수거')
          .not('proceed_type', 'eq', '직접방문')
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

        // 5) 작업완료: ready_to_ship (2026-10-19 출고대기→작업완료, 송장 여부 무관)
        supabase
          .from('repairs')
          .select('*')
          .eq('status', 'ready_to_ship')
          .order('received_at', { ascending: false })
          .limit(50),

        // 6) 출고완료: shipped + delivered + completed — 2026-11-10 최근 30일(끝난 날 기준)만, 최신순
        //    전엔 shipped_at 내림차순 + 50건 상한 → ① 탭 숫자가 늘 50 ② 출고일 없는 방문 건이 맨 위(빈 값이 먼저)
        //    ③ 검색이 이 50건 안에서만 돼 예전 고객을 못 찾음. 이제 기간으로 자르고, 이전 건은 서버 검색(useRepairShippedSearch)
        //    끝난 날 = delivered_at ?? shipped_at ?? updated_at (아래 JS 에서 다시 걸러 정확히 30일)
        //    재수거 접수 후 재작업 전(recall_booked_at 있고 reworked_at 없음)인 건은 '재수리' 탭 → 아래 JS 에서 제외
        supabase
          .from('repairs')
          .select('*')
          .in('status', SHIPPED_STATUSES)
          .or(`delivered_at.gte.${shippedSinceISO},shipped_at.gte.${shippedSinceISO},updated_at.gte.${shippedSinceISO}`)
          .limit(1000),

        // 7) 재수리: 재수거 접수했으나 아직 재작업 전 (출고완료에 묻히던 건 — 여기서 [재작업 시작])
        supabase
          .from('repairs')
          .select('*')
          .not('recall_booked_at', 'is', null)
          .is('reworked_at', null)
          .order('recall_booked_at', { ascending: false })
          .limit(50),

        // 경과일 3일 이상 미처리 — 수거/방문 예정일이 미래면 제외(대기중이라 지연 아님)는 아래 JS에서
        supabase
          .from('repairs')
          .select('pickup_date, visit_date')
          .in('status', ['intake', 'cost_notified'])
          .lt('updated_at', threeDaysAgoISO)
          .limit(200),
      ]);

      const intake = (intakeRes.data || []) as Repair[];
      const pickupNeeded = (pickupNeededRes.data || []) as Repair[];
      const visitScheduled = (visitScheduledRes.data || []) as Repair[];
      const inboundDirect = (inboundWaitingDirectRes.data || []) as Repair[];
      const inboundPickup = (inboundWaitingPickupRes.data || []) as Repair[];
      const inboundWaiting = [...inboundDirect, ...inboundPickup].sort(
        (a, b) => new Date(b.received_at).getTime() - new Date(a.received_at).getTime()
      );
      const inProgress = (inProgressRes.data || []) as Repair[];
      const readyToShip = (readyToShipRes.data || []) as Repair[];
      const shippedSinceMs = new Date(shippedSinceISO).getTime();
      const shipped = sortByEndedAt(((shippedRes.data || []) as Repair[])
        .filter((r) => !isRecallPending(r) && repairEndedAt(r) >= shippedSinceMs));
      // 재수리: 재수거 접수 후 아직 재작업/재출고 전인 것만.
      //  재수거 이후 이미 재출고/재배달된 건(shipped/delivered_at > recall_booked_at)은 사이클 종료 → 제외.
      //  ([재작업 시작] 버튼을 안 거쳐 reworked_at 이 없어도 자동 정리 — 송채림 사례)
      const recall = ((recallRes.data || []) as Repair[]).filter((r) => {
        const rb = r.recall_booked_at ? new Date(r.recall_booked_at).getTime() : 0;
        const reshipped = r.shipped_at && new Date(r.shipped_at).getTime() > rb;
        const redelivered = r.delivered_at && new Date(r.delivered_at).getTime() > rb;
        return !(reshipped || redelivered);
      });

      // 경과 카운트: 수거요청일(pickup_date)·방문일(visit_date)이 미래면 제외(예정일 대기 = 지연 아님)
      const todayStart = new Date(); todayStart.setHours(0, 0, 0, 0);
      const staleRows = (staleRes.data || []) as { pickup_date: string | null; visit_date: string | null }[];
      const staleCount = staleRows.filter((r) => {
        const futurePickup = r.pickup_date && new Date(r.pickup_date).getTime() > todayStart.getTime();
        const futureVisit = r.visit_date && new Date(r.visit_date).getTime() > todayStart.getTime();
        return !(futurePickup || futureVisit);
      }).length;

      return {
        tabData: {
          intake,
          pickup_needed: pickupNeeded,
          visit_scheduled: visitScheduled,
          inbound_waiting: inboundWaiting,
          in_progress: inProgress,
          ready_to_ship: readyToShip,
          shipped,
          recall,
        } as RepairTabData,
        staleCount,
      };
    },
  });

  const tabData = query.data?.tabData || {
    intake: [],
    pickup_needed: [],
    visit_scheduled: [],
    inbound_waiting: [],
    in_progress: [],
    ready_to_ship: [],
    shipped: [],
    recall: [],
  };

  const tabs: RepairTabDef[] = [
    { key: 'intake', label: '신규접수', count: tabData.intake.length },
    { key: 'pickup_needed', label: '수거접수필요', count: tabData.pickup_needed.length },
    { key: 'visit_scheduled', label: '방문예정', count: tabData.visit_scheduled.length },
    { key: 'inbound_waiting', label: '입고대기', count: tabData.inbound_waiting.length },
    { key: 'in_progress', label: '진행중', count: tabData.in_progress.length },
    { key: 'ready_to_ship', label: '작업완료', count: tabData.ready_to_ship.length },
    { key: 'shipped', label: '출고완료', count: tabData.shipped.length },
    { key: 'recall', label: '재수리', count: tabData.recall.length },
  ];

  return {
    tabs,
    tabData,
    isLoading: query.isLoading,
    staleCount: query.data?.staleCount || 0,
  };
}
