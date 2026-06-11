'use client';

import { useState, useEffect } from 'react';
import { useAvailableSerials } from '@/hooks/use-serials';
import { Hash, Check, ChevronDown, ChevronUp } from 'lucide-react';
import { useSerialConflictPrompt } from './serial-conflict-dialog';

interface Props {
  productId: string;
  quantity: number;
  selectedSerialIds: string[];
  onSelect: (serialIds: string[]) => void;
  manualSerials?: string[];
  onManualSerialsChange?: (serials: string[]) => void;
  /** 수정 모드에서 *이미 이 판매에 등록된* 시리얼 (sold 상태) — 해제 가능하게 표시 (2026-05-18 fix) */
  currentSerials?: Array<{ id: string; serial_number: string }>;
  /** 현재 편집 중인 판매 id — 중복 검증 시 같은 판매 시리얼은 충돌로 보지 않기 위함 (Phase A 2026-05-18) */
  currentSaleId?: string;
  /** 사장님이 중복 시리얼 이전 동의 시 호출 — 부모에서 allow_serial_transfer 플래그 set */
  onTransferConsent?: () => void;
  /** 같은 화면(카트)의 다른 품목에 이미 배정된 시리얼 — 자동생성 시 회피해 중복 방지 (2026-06-11 fix) */
  reservedSerials?: string[];
}

export function SerialPicker({ productId, quantity, selectedSerialIds, onSelect, manualSerials = [], onManualSerialsChange, currentSerials = [], currentSaleId, onTransferConsent, reservedSerials = [] }: Props) {
  // Phase A 모달 — Promise 기반 prompt + dialog 노드 (2026-05-18)
  const { prompt: promptConflict, dialog: conflictDialog } = useSerialConflictPrompt();

  /** 시리얼 추가 전 중복 검증 — 충돌 시 사장님 모달 동의 받음 */
  async function confirmIfDuplicate(serial: string): Promise<boolean> {
    try {
      const params = new URLSearchParams({ serial });
      if (currentSaleId) params.set('excludeSaleId', currentSaleId);
      const res = await fetch(`/api/serials/check-duplicate?${params}`);
      const data = await res.json();
      if (!data.exists) return true;
      const ok = await promptConflict({
        serial,
        sale_number: data.sale_number,
        customer_name: data.customer_name,
        product_name: data.product_name,
        sale_date: data.sale_date,
        status: data.status,
      });
      if (ok) onTransferConsent?.(); // 명시 동의 → 부모 플래그 set
      return ok;
    } catch {
      return true; // API 실패 시 보수적 허용
    }
  }

  const [open, setOpen] = useState(false);
  const [manualMode, setManualMode] = useState(false);
  const [manualInput, setManualInput] = useState('');
  const [generating, setGenerating] = useState(false);
  const { data: serials = [], isLoading } = useAvailableSerials(productId);

  // 재고 시리얼이 0개이면 자동으로 직접입력 모드
  useEffect(() => {
    if (!isLoading && serials.length === 0 && open) {
      setManualMode(true);
    }
  }, [isLoading, serials.length, open]);

  function toggleSerial(serialId: string) {
    if (selectedSerialIds.includes(serialId)) {
      onSelect(selectedSerialIds.filter((id) => id !== serialId));
    } else {
      if (selectedSerialIds.length >= quantity) return;
      onSelect([...selectedSerialIds, serialId]);
    }
  }

  async function autoGenerate() {
    // 중복 호출 방지 — 빠른 다중 클릭 시 같은 번호 추가되던 버그 차단 (2026-05-17 fix)
    if (generating) return;
    setGenerating(true);
    try {
      const res = await fetch('/api/serials/batch');
      const data = await res.json();
      let next = data.next_start || 13790001;
      // 이 품목 + 카트 내 다른 품목에 이미 배정된 번호를 모두 회피 (저장 전이라 DB next_start가 동일하게 나오는 문제 방지)
      const existing = [...manualSerials, ...reservedSerials].map((s) => parseInt(s, 10)).filter((n) => !isNaN(n));
      while (existing.includes(next)) next++;
      // Phase A — 자동 생성된 번호도 DB 중복 가능성 (이전 등록·재고 시리얼 등) → 검증 (2026-05-18)
      const ok = await confirmIfDuplicate(String(next));
      if (!ok) return;
      onManualSerialsChange?.([...manualSerials, String(next)]);
    } catch { /* ignore */ }
    finally {
      setGenerating(false);
    }
  }

  /** 직접 입력된 시리얼 추가 (Enter / 추가 버튼) — 중복 검증 후 진행 */
  async function addManualSerial(value: string) {
    const trimmed = value.trim();
    if (!trimmed) return;
    const ok = await confirmIfDuplicate(trimmed);
    if (!ok) return;
    onManualSerialsChange?.([...manualSerials, trimmed]);
    setManualInput('');
  }

  const selectedCount = selectedSerialIds.length + manualSerials.length;
  const isComplete = selectedCount >= quantity;

  return (
    <div className="mt-1">
      {conflictDialog}
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
        <div className="mt-1 p-2 rounded border border-neutral-200 bg-white max-h-64 overflow-y-auto">
          {/* 현재 이 판매에 등록된 시리얼 (sold 상태) — 해제 가능 (2026-05-18 fix) */}
          {currentSerials.length > 0 && (
            <div className="mb-2 pb-2 border-b border-neutral-100">
              <div className="text-[10px] font-semibold text-neutral-500 mb-1 px-1">현재 등록된 시리얼</div>
              <div className="space-y-0.5">
                {currentSerials.map((cs) => {
                  const stillSelected = selectedSerialIds.includes(cs.id);
                  if (!stillSelected) return null; // 해제된 건 표시 안 함
                  return (
                    <div key={cs.id} className="w-full flex items-center gap-2 px-2 py-1.5 rounded text-xs bg-blue-50 text-blue-700">
                      <Check size={10} className="shrink-0" />
                      <span className="font-mono flex-1 truncate">{cs.serial_number}</span>
                      <button
                        type="button"
                        onClick={() => onSelect(selectedSerialIds.filter((id) => id !== cs.id))}
                        className="shrink-0 text-red-400 hover:text-red-600 text-sm leading-none px-1"
                        title="이 판매에서 해제"
                      >×</button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

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
                            ? 'bg-stone-100 text-stone-900 font-semibold'
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
              {/* 자동 번호 생성 — 최상단 눈에 띄는 버튼 (중복 클릭 방지 disabled) */}
              <button type="button" onClick={autoGenerate} disabled={generating}
                className="w-full py-2 rounded-lg border border-dashed border-green-300 bg-green-50 text-green-700 text-xs font-semibold hover:bg-green-100 transition disabled:opacity-50 disabled:cursor-not-allowed">
                {generating ? '생성 중...' : '+ 자동 번호 생성'}
              </button>

              {/* 이미 추가된 시리얼 목록 */}
              {manualSerials.map((s, i) => (
                <div key={i} className="flex items-center gap-1.5">
                  <span className="font-mono text-xs text-neutral-700 flex-1">{s}</span>
                  <button type="button" onClick={() => onManualSerialsChange?.(manualSerials.filter((_, j) => j !== i))}
                    className="text-red-400 hover:text-red-600 text-xs">×</button>
                </div>
              ))}

              {/* 수동 입력 필드 (Phase A — 추가 전 중복 검증) */}
              <div className="flex gap-1.5">
                <input
                  type="text"
                  value={manualInput}
                  onChange={(e) => setManualInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && manualInput.trim()) {
                      e.preventDefault();
                      addManualSerial(manualInput);
                    }
                  }}
                  placeholder="직접 입력 후 엔터"
                  className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs font-mono placeholder:text-neutral-400"
                />
                <button type="button"
                  disabled={!manualInput.trim()}
                  onClick={() => addManualSerial(manualInput)}
                  className="px-2 py-1 text-xs bg-neutral-900 text-white rounded disabled:opacity-30 disabled:cursor-not-allowed">추가</button>
              </div>
            </div>
          )}

          {/* 모드 전환 */}
          <div className="flex items-center gap-2 pt-1.5 mt-1 border-t border-neutral-100">
            <button type="button" onClick={() => setManualMode(!manualMode)}
              className="text-[10px] text-blue-500 hover:text-blue-700">
              {manualMode ? '← 등록된 시리얼' : '직접 입력 →'}
            </button>
          </div>
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
