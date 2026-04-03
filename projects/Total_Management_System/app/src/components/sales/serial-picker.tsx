'use client';

import { useState } from 'react';
import { useAvailableSerials } from '@/hooks/use-serials';
import { Hash, Check, ChevronDown, ChevronUp } from 'lucide-react';

interface Props {
  productId: string;
  quantity: number;
  selectedSerialIds: string[];
  onSelect: (serialIds: string[]) => void;
  manualSerials?: string[];
  onManualSerialsChange?: (serials: string[]) => void;
}

export function SerialPicker({ productId, quantity, selectedSerialIds, onSelect, manualSerials = [], onManualSerialsChange }: Props) {
  const [open, setOpen] = useState(false);
  const [manualMode, setManualMode] = useState(false);
  const [manualInput, setManualInput] = useState('');
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
        <div className="mt-1 p-2 rounded border border-neutral-200 bg-white max-h-40 overflow-y-auto">
          {/* 등록된 시리얼 선택 */}
          {!manualMode && (
            <>
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
            </>
          )}

          {/* 직접 입력 모드 */}
          {manualMode && (
            <div className="space-y-1.5">
              {manualSerials.map((s, i) => (
                <div key={i} className="flex items-center gap-1.5">
                  <span className="font-mono text-xs text-neutral-700 flex-1">{s}</span>
                  <button type="button" onClick={() => onManualSerialsChange?.(manualSerials.filter((_, j) => j !== i))}
                    className="text-red-400 hover:text-red-600 text-xs">×</button>
                </div>
              ))}
              <div className="flex gap-1.5">
                <input
                  type="text"
                  value={manualInput}
                  onChange={(e) => setManualInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && manualInput.trim()) {
                      e.preventDefault();
                      onManualSerialsChange?.([...manualSerials, manualInput.trim()]);
                      setManualInput('');
                    }
                  }}
                  placeholder="시리얼 번호 입력 후 엔터"
                  className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs font-mono placeholder:text-neutral-400"
                />
                <button type="button"
                  onClick={() => { if (manualInput.trim()) { onManualSerialsChange?.([...manualSerials, manualInput.trim()]); setManualInput(''); } }}
                  className="px-2 py-1 text-xs bg-neutral-900 text-white rounded">추가</button>
              </div>
            </div>
          )}

          {/* 모드 전환 */}
          <button
            type="button"
            onClick={() => setManualMode(!manualMode)}
            className="w-full text-center text-[10px] text-blue-500 hover:text-blue-700 pt-1.5 mt-1 border-t border-neutral-100"
          >
            {manualMode ? '← 등록된 시리얼에서 선택' : '시리얼 직접 입력 →'}
          </button>
        </div>
      )}

      {/* 직접 입력된 시리얼 표시 (닫혀있을 때) */}
      {!open && manualSerials.length > 0 && (
        <div className="mt-0.5 flex flex-wrap gap-1">
          {manualSerials.map((s, i) => (
            <span key={i} className="px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 text-[10px] font-mono">{s}</span>
          ))}
        </div>
      )}
    </div>
  );
}
