'use client';

import { Badge } from '@/components/ui/badge';
import { REPAIR_STATUS_LABEL, REPAIR_STATUS_COLOR } from '@/lib/utils/format';

interface RepairStatusBadgeProps {
  status: string;
  className?: string;
}

export function RepairStatusBadge({ status, className = '' }: RepairStatusBadgeProps) {
  const color = REPAIR_STATUS_COLOR[status] || 'bg-neutral-100 text-neutral-500';
  const label = REPAIR_STATUS_LABEL[status] || status;

  return (
    <Badge className={`${color} ${className}`}>
      {label}
    </Badge>
  );
}
