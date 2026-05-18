'use client';

import { useCallback, useState } from 'react';
import { AlertTriangle, X } from 'lucide-react';

export interface SerialConflictInfo {
  serial: string;
  sale_number: string | null;
  customer_name: string | null;
  product_name: string | null;
  sale_date: string | null;
  status: string | null;
}

interface DialogProps {
  info: SerialConflictInfo;
  onConfirm: () => void;
  onCancel: () => void;
}

/** 시리얼 중복 충돌 다이얼로그 (Brand Guide 모노크롬) */
export function SerialConflictDialog({ info, onConfirm, onCancel }: DialogProps) {
  return (
    <div
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 backdrop-blur-sm p-4"
      onClick={onCancel}
    >
      <div
        className="bg-white rounded-xl shadow-2xl w-full max-w-[420px] overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="px-5 py-4 border-b border-neutral-100 flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-full bg-amber-50 flex items-center justify-center flex-shrink-0">
            <AlertTriangle size={16} className="text-amber-600" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-sm font-bold text-neutral-900 leading-tight">시리얼 중복 확인</h3>
            <p className="text-[11px] text-neutral-500 mt-0.5 font-mono truncate">#{info.serial}</p>
          </div>
          <button
            onClick={onCancel}
            className="w-7 h-7 rounded-full hover:bg-neutral-100 flex items-center justify-center text-neutral-400 hover:text-neutral-700 transition"
            aria-label="닫기"
          >
            <X size={14} />
          </button>
        </div>

        {/* 본문 */}
        <div className="px-5 py-4 space-y-3">
          <p className="text-[13px] text-neutral-700 leading-relaxed">
            이 시리얼은 <strong>이미 다른 판매에 등록</strong>되어 있습니다.
          </p>

          <div className="bg-neutral-50 rounded-lg p-3 space-y-1.5">
            <Row label="판매번호" value={info.sale_number} mono />
            <Row label="고객" value={info.customer_name} />
            <Row label="제품" value={info.product_name} />
            <Row label="판매일" value={info.sale_date} />
            <Row label="상태" value={info.status === 'sold' ? '판매완료' : info.status} />
          </div>

          <div className="text-[12px] text-neutral-500 leading-relaxed border-t border-neutral-100 pt-3">
            <p className="mb-1">
              <strong className="text-neutral-900">이전 판매에서 분리</strong>하여 이쪽으로 가져옵니다.
            </p>
            <p className="text-amber-700">이전 판매의 시리얼은 사라집니다.</p>
          </div>
        </div>

        {/* 푸터 */}
        <div className="px-5 py-3 border-t border-neutral-100 flex gap-2 bg-neutral-50/50">
          <button
            onClick={onCancel}
            className="flex-1 h-9 rounded-lg border border-neutral-200 bg-white text-sm font-medium text-neutral-700 hover:bg-neutral-50 transition"
          >
            취소
          </button>
          <button
            onClick={onConfirm}
            className="flex-1 h-9 rounded-lg bg-neutral-900 text-sm font-semibold text-white hover:bg-neutral-800 transition"
          >
            이전 판매에서 가져오기
          </button>
        </div>
      </div>
    </div>
  );
}

function Row({ label, value, mono }: { label: string; value: string | null; mono?: boolean }) {
  return (
    <div className="flex items-baseline gap-2 text-[12px]">
      <span className="text-neutral-500 w-12 flex-shrink-0">{label}</span>
      <span className={`text-neutral-900 font-medium ${mono ? 'font-mono' : ''}`}>{value || '-'}</span>
    </div>
  );
}

/**
 * Promise 기반 시리얼 충돌 다이얼로그 hook.
 * - prompt(info) → Promise<boolean> 반환 (확인 true / 취소 false)
 * - dialog 노드를 컴포넌트 트리에 같이 렌더링
 */
export function useSerialConflictPrompt() {
  const [state, setState] = useState<{ info: SerialConflictInfo; resolve: (v: boolean) => void } | null>(null);

  const prompt = useCallback((info: SerialConflictInfo): Promise<boolean> => {
    return new Promise<boolean>((resolve) => setState({ info, resolve }));
  }, []);

  const dialog = state ? (
    <SerialConflictDialog
      info={state.info}
      onConfirm={() => {
        state.resolve(true);
        setState(null);
      }}
      onCancel={() => {
        state.resolve(false);
        setState(null);
      }}
    />
  ) : null;

  return { prompt, dialog };
}
