'use client';

import { Check, X, Award, RotateCcw } from 'lucide-react';
import type { DemoPOApi } from './use-demo-po';
import type { DemoPOItem } from './types';
import { STATUS_LABEL, STATUS_TONE } from './types';

/**
 * PC 그리드 매칭 현황 뷰.
 * 모바일 매칭 화면과 같은 useState 를 공유 → 모바일에서 매칭하면 여기 셀이 emerald로 즉시 전환.
 */
export function InboundGridPC({
  api,
  onPromote,
}: {
  api: DemoPOApi;
  onPromote: (itemId: string) => void;
}) {
  const { po, totals, select, selectedItemId, reject, setStatus } = api;

  return (
    <div className="space-y-4">
      {/* 진행 현황 바 */}
      <div className="rounded-xl bg-white border border-stone-200 p-4">
        <div className="flex items-center justify-between flex-wrap gap-3 mb-3">
          <div className="text-sm font-bold text-stone-900">매칭 현황</div>
          <div className="text-xs text-stone-500">
            전체 {po.items.length}건
          </div>
        </div>
        <div className="grid grid-cols-4 gap-2">
          <Stat label="대기" value={totals.pending} tone="bg-stone-100 text-stone-700" />
          <Stat label="매칭" value={totals.matched} tone="bg-emerald-50 text-emerald-700" />
          <Stat label="정식등록" value={totals.promoted} tone="bg-stone-900 text-white" />
          <Stat label="보류" value={totals.rejected} tone="bg-rose-50 text-rose-600" />
        </div>
      </div>

      {/* 그리드 */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
        {po.items.map((it) => (
          <GridCell
            key={it.id}
            item={it}
            selected={selectedItemId === it.id}
            onClick={() => select(it.id)}
            onPromote={() => onPromote(it.id)}
            onReject={() => reject(it.id)}
            onReset={() => setStatus(it.id, 'pending')}
          />
        ))}
      </div>

      <div className="text-[11px] text-stone-400 italic">
        ※ 셀 클릭 → 좌측 모바일 화면에 해당 품목이 열립니다. 사진 매칭하면 셀이 emerald로 변경.
      </div>
    </div>
  );
}

function Stat({ label, value, tone }: { label: string; value: number; tone: string }) {
  return (
    <div className={`rounded-lg px-3 py-2 ${tone}`}>
      <div className="text-[10px] uppercase tracking-wider opacity-70">{label}</div>
      <div className="text-xl font-bold leading-tight">{value}</div>
    </div>
  );
}

function GridCell({
  item,
  selected,
  onClick,
  onPromote,
  onReject,
  onReset,
}: {
  item: DemoPOItem;
  selected: boolean;
  onClick: () => void;
  onPromote: () => void;
  onReject: () => void;
  onReset: () => void;
}) {
  const seq = item.sticker_no.split('-').pop();
  const firstPhoto = item.inbound_photos[0];
  const status = item.inspection_status;

  const borderCls = selected
    ? 'border-stone-900 ring-2 ring-stone-900/20'
    : status === 'matched'
      ? 'border-emerald-300'
      : status === 'promoted'
        ? 'border-stone-900 bg-stone-50'
        : status === 'rejected'
          ? 'border-rose-200 bg-rose-50/30 opacity-60'
          : 'border-dashed border-stone-300 bg-stone-50';

  return (
    <div
      className={`group relative rounded-xl border-2 overflow-hidden transition ${borderCls}`}
    >
      <button
        type="button"
        onClick={onClick}
        className="block w-full aspect-square relative text-left"
      >
        {/* 배경: 사진 또는 그라데이션 */}
        {firstPhoto ? (
          <div
            className="absolute inset-0"
            style={{
              background: firstPhoto.startsWith('mock:')
                ? firstPhoto.split(':').slice(2).join(':')
                : `url(${firstPhoto}) center/cover`,
            }}
          />
        ) : (
          <div className="absolute inset-0 flex items-center justify-center text-5xl text-stone-300 font-bold">
            ?
          </div>
        )}

        {/* 상단: 번호 + 상태 뱃지 */}
        <div className="absolute top-2 left-2 right-2 flex items-start justify-between gap-1">
          <span className="inline-flex items-center justify-center min-w-[28px] h-7 px-2 rounded-full bg-white/95 text-stone-900 text-xs font-bold shadow-sm">
            {seq}
          </span>
          <span
            className={`inline-flex items-center px-2 py-0.5 rounded-full text-[10px] border font-medium ${STATUS_TONE[status]} backdrop-blur-sm`}
          >
            {STATUS_LABEL[status]}
          </span>
        </div>

        {/* 하단: 품목명 */}
        <div className="absolute bottom-0 left-0 right-0 p-2 bg-gradient-to-t from-black/70 to-transparent">
          <div className="text-[11px] font-bold text-white truncate">
            {item.product_name || '(품목명 없음)'}
          </div>
          {item.promoted_sku && (
            <div className="text-[10px] text-amber-200 font-mono mt-0.5">
              {item.promoted_sku}
            </div>
          )}
        </div>
      </button>

      {/* 호버 액션 (matched/rejected/promoted 단계별) */}
      <div className="absolute inset-x-0 bottom-0 opacity-0 group-hover:opacity-100 transition pointer-events-none">
        <div className="px-2 pb-2 flex items-center gap-1 justify-end pointer-events-auto">
          {status === 'matched' && (
            <>
              <button
                type="button"
                onClick={onPromote}
                className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-stone-900 text-white text-[10px] font-bold hover:bg-stone-700"
              >
                <Award size={11} /> 정식 채택
              </button>
              <button
                type="button"
                onClick={onReject}
                className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-white text-stone-700 text-[10px] font-medium border border-stone-200 hover:bg-stone-50"
              >
                <X size={11} /> 보류
              </button>
            </>
          )}
          {status === 'pending' && firstPhoto && (
            <button
              type="button"
              onClick={() => {
                // pending인데 사진이 있으면 강제 matched 처리 (시뮬 편의)
                onClick();
              }}
              className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-emerald-600 text-white text-[10px] font-bold hover:bg-emerald-700"
            >
              <Check size={11} /> 매칭완료로
            </button>
          )}
          {status === 'rejected' && (
            <button
              type="button"
              onClick={onReset}
              className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-white text-stone-700 text-[10px] font-medium border border-stone-200 hover:bg-stone-50"
            >
              <RotateCcw size={11} /> 되돌리기
            </button>
          )}
        </div>
      </div>
    </div>
  );
}
