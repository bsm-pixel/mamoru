'use client';

import { useLocations, useAssignLocation } from '@/hooks/use-warehouse';
import { MapPin } from 'lucide-react';

/**
 * 제품 정위치 지정 — 112, 2026-07-18
 * 드롭다운으로 자리를 고르면 products.location_id 만 바뀐다 (재고 수량 무관).
 */
export function LocationPicker({ productId, currentLocationId }: {
  productId: string;
  currentLocationId: string | null;
}) {
  const { data, isLoading } = useLocations();
  const assign = useAssignLocation();
  const locations = data?.locations || [];

  if (isLoading) {
    return <div className="h-9 rounded-lg bg-neutral-100 animate-pulse" />;
  }

  if (locations.length === 0) {
    return (
      <p className="text-[11px] text-neutral-400">
        등록된 자리가 없습니다 — 창고 배치도에서 렉을 먼저 만들어주세요
      </p>
    );
  }

  return (
    <div className="flex items-center gap-2">
      <MapPin size={14} className="text-neutral-400 shrink-0" />
      <select
        value={currentLocationId || ''}
        disabled={assign.isPending}
        onChange={(e) => assign.mutate({ product_ids: [productId], location_id: e.target.value || null })}
        className="flex-1 h-9 px-2 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-neutral-300 disabled:opacity-50"
      >
        <option value="">위치 미지정</option>
        {locations.map((l) => (
          <option key={l.id} value={l.id}>
            {l.code} · {l.label || ''}{l.product_count > 0 ? ` (${l.product_count}종)` : ''}
          </option>
        ))}
      </select>
    </div>
  );
}
