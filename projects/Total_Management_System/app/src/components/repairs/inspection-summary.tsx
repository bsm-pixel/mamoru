'use client';

import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { generateWorkSummary, getScissorSummary } from '@/lib/repair/inspection-text';
import type { RepairInspection } from '@/lib/supabase/types';
import { ClipboardCheck } from 'lucide-react';

interface InspectionSummaryProps {
  inspections: RepairInspection[];
}

export function InspectionSummary({ inspections }: InspectionSummaryProps) {
  if (inspections.length === 0) {
    return (
      <Card>
        <CardHeader>
          <CardTitle>
            <ClipboardCheck size={16} className="inline mr-1.5" />
            검수 결과
          </CardTitle>
        </CardHeader>
        <p className="text-sm text-neutral-400">검수 데이터 없음</p>
      </Card>
    );
  }

  const workSummary = generateWorkSummary(inspections);

  return (
    <Card>
      <CardHeader>
        <CardTitle>
          <ClipboardCheck size={16} className="inline mr-1.5" />
          검수 결과
        </CardTitle>
      </CardHeader>

      {/* 가위별 요약 */}
      <div className="space-y-2 mb-4">
        {inspections.map((insp) => (
          <div key={insp.id} className="flex items-start gap-2 text-sm">
            <span className="font-mono text-xs bg-neutral-100 px-2 py-0.5 rounded shrink-0">
              #{insp.scissor_number}
            </span>
            <div>
              <span className="text-neutral-500 mr-1">{insp.scissor_type || '기타'}</span>
              <span className="text-neutral-700">{getScissorSummary(insp)}</span>
            </div>
          </div>
        ))}
      </div>

      {/* 종합 문구 */}
      {workSummary && (
        <div className="border-t border-neutral-100 pt-3">
          <p className="text-xs text-neutral-500 mb-1">자동 생성 수리내역</p>
          <p className="text-sm text-neutral-700 whitespace-pre-wrap bg-neutral-50 p-3 rounded-lg">
            {workSummary}
          </p>
        </div>
      )}
    </Card>
  );
}
