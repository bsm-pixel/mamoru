'use client';

import { useState } from 'react';
import { useAvailableSerials } from '@/hooks/use-serials';
import { Hash, Check, ChevronDown, ChevronUp } from 'lucide-react';

interface Props {
  productId: string;
  quantity: number;
  selectedSerialIds: string[];
  onSelect: (serialIds: string[]) => void;
}

export function SerialPicker({ productId, quantity, selectedSerialIds, onSelect }: Props) {
  const [open, setOpen] = useState(false);
  const { data: serials = [], isLoading } = useAvailableSerials(productId);

  function toggleSerial(serialId: string) {
    if (selectedSerialIds.includes(serialId)) {
      onSelect(selectedSerialIds.filter((id) => id !== serialId));
    } else {
      // 수량 이상 선택 방지
      if (selectedSerialIds.length >= quantity) return;
      onSelect([...selectedSerialIds, serialId]);
    }
  }

  const selectedCount = selectedSerialIds.length;
  const isComplete = selectedCount === quantity;

  return (
    <div className="mt-1">
      <button
        type="button"
        onClick={() => setOpen(!open)}
        className={`flex items-center gap-1.5 text-xs px-2 py-1 rounded transition ${
          isComplete
            ? 'bg-green-50 text-green-700'
            : selectedCount > 0
              ? 'bg-yellow-50 text-yellow-700'
              : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
        }`}
      >
        <Hash size={10} />
        <span>
          시리얼 {selectedCount}/{quantity}
          {serials.length > 0 && ` (재고 ${serials.length})`}
        </span>
        {open ? <ChevronUp size={10} /> : <ChevronDown size={10} />}
      </button>

      {open && (
        <div className="mt-1 p-2 rounded border border-neutral-200 bg-white max-h-32 overflow-y-auto">
          {isLoading ? (
            <p className="text-xs text-neutral-400 text-center py-2">로딩중...</p>
          ) : serials.length === 0 ? (
            <p className="text-xs text-neutral-400 text-center py-2">재고 시리얼 없음</p>
          ) : (
            <div className="space-y-0.5">
              {serials.map((s) => {
                const isSelected = selectedSerialIds.includes(s.id);
                const isDisabled = !isSelected && selectedSerialIds.length >= quantity;
                return (
                  <button
                    key={s.id}
                    type="button"
                    onClick={() => toggleSerial(s.id)}
                    disabled={isDisabled}
                    className={`w-full flex items-center gap-2 px-2 py-1.5 rounded text-left text-xs transition ${
                      isSelected
                        ? 'bg-terracotta/10 text-terracotta font-semibold'
                        : isDisabled
                          ? 'text-neutral-300 cursor-not-allowed'
                          : 'hover:bg-neutral-50 text-neutral-700'
                    }`}
                  >
                    {isSelected && <Check size={10} className="shrink-0" />}
                    <span className="font-mono truncate">{s.serial_number}</span>
                    {s.warehouse_zone === 'display' && (
                      <span className="shrink-0 px-1 py-0.5 rounded text-[9px] font-semibold bg-purple-50 text-purple-600">디스플레이</span>
                    )}
                  </button>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
