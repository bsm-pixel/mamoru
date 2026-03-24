'use client';

import { useDeliveryTrack } from '@/hooks/use-delivery-track';
import { Skeleton } from '@/components/ui/skeleton';
import { ExternalLink, XCircle, HelpCircle } from 'lucide-react';

const STEPS = [
  { label: '접수', codes: ['01'] },
  { label: '집화', codes: ['02', '41'] },
  { label: '배달중', codes: ['42', '44'] },
  { label: '배달완료', codes: ['91'] },
];

function getActiveStep(tracking: Array<{ godsStatCd?: string }>) {
  let maxStep = -1;
  for (const t of tracking) {
    const code = String(t.godsStatCd || '');
    for (let i = 0; i < STEPS.length; i++) {
      if (STEPS[i].codes.includes(code) && i > maxStep) maxStep = i;
    }
  }
  return maxStep;
}

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

  const tracking = data.raw?.tracking || [];
  const activeStep = data.state === 'DELIVERED' ? 3 : getActiveStep(tracking);

  // 마지막 추적 정보
  const lastTrack = tracking.length > 0 ? tracking[tracking.length - 1] : null;

  return (
    <div className="space-y-3">
      {/* 4단계 스텝 인디케이터 */}
      <div className="flex items-center gap-0">
        {STEPS.map((step, i) => (
          <div key={step.label} className="flex items-center flex-1">
            <div className="flex flex-col items-center flex-1">
              <div className={`w-6 h-6 rounded-full flex items-center justify-center text-[10px] font-bold transition ${
                i <= activeStep
                  ? 'bg-terracotta text-white'
                  : 'bg-neutral-200 text-neutral-400'
              }`}>
                {i + 1}
              </div>
              <span className={`mt-1 text-[10px] font-medium ${
                i <= activeStep ? 'text-terracotta' : 'text-neutral-400'
              }`}>
                {step.label}
              </span>
            </div>
            {i < STEPS.length - 1 && (
              <div className={`h-0.5 flex-1 -mx-1 ${
                i < activeStep ? 'bg-terracotta' : 'bg-neutral-200'
              }`} />
            )}
          </div>
        ))}
      </div>

      {/* 마지막 추적 정보 */}
      {lastTrack && (
        <p className="text-xs text-neutral-500">
          {lastTrack.statDt} {lastTrack.statTm} {lastTrack.orgNm && `· ${lastTrack.orgNm}`}
        </p>
      )}

      {/* 외부 추적 링크 */}
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
