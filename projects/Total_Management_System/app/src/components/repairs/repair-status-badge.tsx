'use client';

import { Badge } from '@/components/ui/badge';
import { REPAIR_STATUS_COLOR } from '@/lib/utils/format';
import { getRepairDisplayLabel } from '@/lib/repair/transitions';
import type { RepairStatus } from '@/lib/supabase/types';

interface RepairStatusBadgeProps {
  status: string;
  proceedType?: string | null;
  className?: string;
}

export function RepairStatusBadge({ status, proceedType, className = '' }: RepairStatusBadgeProps) {
  const color = REPAIR_STATUS_COLOR[status] || 'bg-neutral-100 text-neutral-500';
  const label = getRepairDisplayLabel(status as RepairStatus, proceedType);

  return (
    <Badge className={`${color} ${className}`}>
      {label}
    </Badge>
  );
}
