'use client';

import { Card } from '@/components/ui/card';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairActionChips } from '../repair-action-chips';
import { RepairStatusBadge } from '../repair-status-badge';
import { formatPhone, formatKRW } from '@/lib/utils/format';
import type { Repair } from '@/lib/supabase/types';
import { Scissors } from 'lucide-react';
import Link from 'next/link';

interface InProgressTabProps {
  repairs: Repair[];
  isLoading: boolean;
}

/** 진행중 탭: cost_notified + repairing — 인라인 칩 바 (내역서/입금/송장/포장) */
export function InProgressTab({ repairs, isLoading }: InProgressTabProps) {
  if (isLoading) {
    return (
      <div className="space-y-3 p-4">
        {[...Array(3)].map((_, i) => <Skeleton key={i} className="h-32 rounded-xl" />)}
      </div>
    );
  }

  if (!repairs.length) {
    return (
      <div className="flex flex-col items-center justify-center py-16 text-neutral-400">
        <Scissors size={32} className="mb-2 opacity-50" />
        <p className="text-sm">진행중인 복원수리 건이 없습니다</p>
      </div>
    );
  }

  return (
    <div className="space-y-2 p-4">
      {repairs.map((r) => (
        <Card key={r.id} className="hover:bg-neutral-50 transition">
          <Link href={`/repairs/${r.id}`} className="block">
            <div className="flex items-start justify-between gap-3">
              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2 mb-1">
                  <span className="text-sm font-semibold">{r.name}</span>
                  <span className="text-xs text-neutral-400 font-mono">{r.as_id}</span>
                  <RepairStatusBadge status={r.status} />
                </div>
                <div className="flex items-center gap-3 text-xs text-neutral-600">
                  {r.qty_mamoru > 0 && <span>마모루 {r.qty_mamoru}자루</span>}
                  {r.qty_other > 0 && <span>타사 {r.qty_other}자루</span>}
                </div>
              </div>
              {r.total_amount > 0 && (
                <span className="text-sm font-bold text-terracotta-deep shrink-0">
                  {formatKRW(r.total_amount)}
                </span>
              )}
            </div>
          </Link>
          {/* 인라인 칩 바 — Link 바깥에서 클릭 가능 */}
          <RepairActionChips repair={r} />
        </Card>
      ))}
    </div>
  );
}
