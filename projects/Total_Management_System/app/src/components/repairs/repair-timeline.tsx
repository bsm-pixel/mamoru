'use client';

import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { REPAIR_STATUS_LABEL } from '@/lib/utils/format';
import { formatRelative } from '@/lib/utils/format';
import type { RepairHistory } from '@/lib/supabase/types';
import { Clock } from 'lucide-react';

interface RepairTimelineProps {
  history: RepairHistory[];
}

export function RepairTimeline({ history }: RepairTimelineProps) {
  return (
    <Card>
      <CardHeader>
        <CardTitle>
          <Clock size={16} className="inline mr-1.5" />
          상태 이력
        </CardTitle>
      </CardHeader>
      {history.length === 0 ? (
        <p className="text-sm text-neutral-400">이력 없음</p>
      ) : (
        <div className="space-y-3">
          {history.map((h) => (
            <div key={h.id} className="flex items-start gap-3 text-sm">
              <div className="w-2 h-2 rounded-full bg-terracotta mt-1.5 shrink-0" />
              <div>
                <div className="flex items-center gap-2">
                  {h.from_status && (
                    <>
                      <span className="text-neutral-500">
                        {REPAIR_STATUS_LABEL[h.from_status] || h.from_status}
                      </span>
                      <span className="text-neutral-400">&rarr;</span>
                    </>
                  )}
                  <span className="font-medium">
                    {REPAIR_STATUS_LABEL[h.to_status] || h.to_status}
                  </span>
                </div>
                {h.note && <p className="text-xs text-neutral-500 mt-0.5">{h.note}</p>}
                <p className="text-xs text-neutral-400 mt-0.5">
                  {formatRelative(h.created_at)}
                </p>
              </div>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}
