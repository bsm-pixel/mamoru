'use client';

import { useDeliveryTrack } from '@/hooks/use-delivery-track';
import { Skeleton } from '@/components/ui/skeleton';
import { StatusStepper } from '@/components/ui/status-stepper';
import { TRACK_STEPS, resolveStepIndex, stepTimes, formatLastScan, type TrackRecord } from '@/lib/lotte/track-steps';
import { ExternalLink, XCircle, HelpCircle } from 'lucide-react';

/**
 * 롯데 배송 추적 스텝바 (판매·주문·복원수리 3곳 공용)
 *
 * 2026-09-26 전면 수정:
 *   · 단계 판정을 실측 코드표(lib/lotte/track-steps.ts)로 이관 — 송장만 출력한 건이 '집하'로 보이던 문제 해결
 *   · 자체 스텝바(테라코타 원) → TMS 공용 StatusStepper 로 통일(완료 ✓ / 현재 검정 / 이후 회색)
 *   · 마지막 스캔 줄이 없는 필드(statDt·statTm·orgNm)를 읽어 빈 줄이던 것 → 실제 필드로 교체
 */
interface Props {
  invNo: string;
}

export function DeliveryTracker({ invNo }: Props) {
  const { data, isLoading } = useDeliveryTrack(invNo);

  if (isLoading) return <Skeleton className="h-16 w-full" />;
  if (!data || !data.ok) {
    if (data?.state === 'NOT_FOUND') {
      return (
        <div className="flex items-center gap-2 p-3 rounded-lg bg-neutral-50 text-sm text-neutral-500">
          <HelpCircle size={16} />
          추적 정보가 아직 없습니다
        </div>
      );
    }
    return null;
  }

  if (data.state === 'CANCELLED') {
    return (
      <div className="flex items-center gap-2 p-3 rounded-lg bg-red-50 text-sm text-red-600">
        <XCircle size={16} />
        배송 취소됨
      </div>
    );
  }

  const tracking = (data.raw?.tracking || []) as TrackRecord[];
  const times = stepTimes(tracking);
  // 배달완료는 state 로도 판정된다(41/45) — 코드가 누락돼도 마지막 단계를 보장
  const idx = data.state === 'DELIVERED' ? TRACK_STEPS.length - 1 : resolveStepIndex(tracking);
  const current = TRACK_STEPS[Math.max(idx, 0)];
  const last = formatLastScan(tracking);

  return (
    <div className="space-y-2">
      <StatusStepper
        steps={TRACK_STEPS.map((s) => ({ key: s.key, label: s.label, at: times[s.key] }))}
        currentKey={current.key}
      />

      {/* 마지막 스캔 — 롯데가 보내는 안내 문장을 그대로 (예: '물품을 받으셨습니다.') */}
      {last && (
        <p className="text-[11px] text-neutral-500 leading-snug">
          {[last.label, last.place, last.atText].filter(Boolean).join(' · ')}
          {last.message && <span className="block text-neutral-400">{last.message}</span>}
        </p>
      )}

      <a
        href={`https://www.lotteglogis.com/home/reservation/tracking/linkView?InvNo=${invNo}`}
        target="_blank"
        rel="noopener noreferrer"
        className="inline-flex items-center gap-1 text-xs text-blue-600 hover:text-blue-700"
      >
        <ExternalLink size={12} />
        롯데택배에서 보기
      </a>
    </div>
  );
}
