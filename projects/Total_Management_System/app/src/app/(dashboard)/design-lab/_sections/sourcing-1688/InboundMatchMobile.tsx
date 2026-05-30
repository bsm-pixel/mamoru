'use client';

import { useState, useRef, useEffect } from 'react';
import { Camera, X, ExternalLink, ArrowLeft, Check } from 'lucide-react';
import type { DemoPOApi } from './use-demo-po';
import { STATUS_LABEL, STATUS_TONE } from './types';

/**
 * 모바일 입고매칭 화면 시뮬레이션.
 * 실제 운영에서는 `/purchasing/inbound/[itemId]/page.tsx` 가 됨.
 * 데모: 카메라 버튼 클릭 시 가짜 사진 1장 추가 (그라데이션 base64 dataURL).
 */

const MOCK_PHOTOS = [
  // 가위 느낌의 회색 그라데이션 (실제 사진 대신)
  'linear-gradient(135deg, #44403c 0%, #78716c 50%, #d6d3d1 100%)',
  'linear-gradient(135deg, #57534e 0%, #a8a29e 50%, #f5f5f4 100%)',
  'linear-gradient(135deg, #292524 0%, #57534e 50%, #a8a29e 100%)',
  'linear-gradient(135deg, #1c1917 0%, #44403c 50%, #78716c 100%)',
  'linear-gradient(135deg, #44403c 0%, #d6d3d1 100%)',
];

function makeMockPhoto(seed: number): string {
  // 실제 base64 이미지 대신 CSS gradient 토큰. 렌더링 단에서 background로 처리
  return `mock:${seed}:${MOCK_PHOTOS[seed % MOCK_PHOTOS.length]}`;
}

export function InboundMatchMobile({ api }: { api: DemoPOApi }) {
  const { po, selectedItemId, select, addPhoto, removePhoto, completeMatch, updateItem } = api;
  const seedRef = useRef(0);

  const item = po.items.find((it) => it.id === selectedItemId);
  const [memo, setMemo] = useState('');

  // 다른 item 선택될 때만 memo 동기화 (렌더 중 setState 금지)
  useEffect(() => {
    if (item) setMemo(item.inbound_memo);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [item?.id]);

  // 미선택 화면 — item 리스트
  if (!item) {
    return <ItemPicker api={api} />;
  }

  const photos = item.inbound_photos;
  const seq = item.sticker_no.split('-').pop();

  const handleCapture = () => {
    const next = makeMockPhoto(seedRef.current++);
    addPhoto(item.id, next);
  };

  const handleComplete = () => {
    completeMatch(item.id, memo);
    // 다음 pending item으로 자동 이동
    const nextPending = po.items.find((it) => it.inspection_status === 'pending' && it.id !== item.id);
    if (nextPending) {
      select(nextPending.id);
    }
  };

  const handleQuantity = (delta: number) => {
    const next = Math.max(1, item.quantity + delta);
    updateItem(item.id, { quantity: next });
  };

  return (
    <div className="min-h-full bg-stone-50">
      {/* 헤더 */}
      <div className="sticky top-0 z-10 bg-white border-b border-stone-200 px-4 py-3 flex items-center gap-3">
        <button
          type="button"
          onClick={() => select(null)}
          className="text-stone-600 active:text-stone-900"
        >
          <ArrowLeft size={20} />
        </button>
        <div className="flex-1">
          <div className="text-[10px] text-stone-400 font-mono">{item.sticker_no}</div>
          <div className="text-sm font-bold text-stone-900">입고매칭 #{seq}</div>
        </div>
        <span
          className={`px-2 py-0.5 text-[10px] rounded-full border font-medium ${STATUS_TONE[item.inspection_status]}`}
        >
          {STATUS_LABEL[item.inspection_status]}
        </span>
      </div>

      <div className="p-4 space-y-3 pb-24">
        {/* 품목 정보 카드 */}
        <div className="rounded-xl bg-white p-4 shadow-sm">
          <div className="text-base font-bold text-stone-900 leading-snug">
            {item.product_name || '(품목명 없음)'}
          </div>
          <div className="mt-2 text-xs text-stone-500 flex items-center flex-wrap gap-x-3 gap-y-1">
            <span>
              단가 <strong className="text-stone-900">¥{item.unit_price}</strong>
            </span>
            <span>·</span>
            <span>
              수량 <strong className="text-stone-900">{item.quantity}</strong>
            </span>
            {item.moq && (
              <>
                <span>·</span>
                <span>MOQ {item.moq}</span>
              </>
            )}
          </div>
          {item.features_memo && (
            <div className="mt-3 px-3 py-2 bg-amber-50 border-l-2 border-amber-300 text-xs text-stone-700 rounded">
              {item.features_memo}
            </div>
          )}
          {item.vendor_url && (
            <a
              href={item.vendor_url}
              target="_blank"
              rel="noreferrer"
              className="mt-3 inline-flex items-center gap-1 text-xs text-blue-600 active:text-blue-800"
            >
              1688 상품 페이지 <ExternalLink size={11} />
            </a>
          )}
        </div>

        {/* 수량 조정 */}
        <div className="rounded-xl bg-white p-4 shadow-sm">
          <div className="text-[11px] font-bold text-stone-500 uppercase tracking-wider mb-2">
            실제 입고 수량
          </div>
          <div className="flex items-center justify-between">
            <span className="text-xs text-stone-500">예상 {item.quantity}개</span>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => handleQuantity(-1)}
                className="w-9 h-9 rounded-lg bg-stone-100 active:bg-stone-200 text-lg font-bold text-stone-700"
              >
                −
              </button>
              <span className="text-xl font-bold text-stone-900 w-10 text-center">
                {item.quantity}
              </span>
              <button
                type="button"
                onClick={() => handleQuantity(1)}
                className="w-9 h-9 rounded-lg bg-stone-100 active:bg-stone-200 text-lg font-bold text-stone-700"
              >
                +
              </button>
            </div>
          </div>
        </div>

        {/* 사진 영역 */}
        <div className="rounded-xl bg-white p-4 shadow-sm">
          <div className="flex items-center justify-between mb-3">
            <div className="text-[11px] font-bold text-stone-500 uppercase tracking-wider">
              현장 사진
            </div>
            <span className="text-xs text-stone-500">{photos.length}/5</span>
          </div>
          <button
            type="button"
            onClick={handleCapture}
            disabled={photos.length >= 5}
            className="w-full py-5 bg-stone-900 active:bg-stone-700 text-white font-bold rounded-xl flex items-center justify-center gap-2 disabled:opacity-40 disabled:cursor-not-allowed"
          >
            <Camera size={20} />
            {photos.length === 0 ? '사진 촬영' : '추가 촬영'}
          </button>

          {photos.length > 0 && (
            <div className="mt-3 grid grid-cols-3 gap-2">
              {photos.map((p, i) => (
                <div
                  key={i}
                  className="relative aspect-square rounded-lg overflow-hidden border border-stone-200"
                  style={{ background: p.startsWith('mock:') ? p.split(':').slice(2).join(':') : `url(${p}) center/cover` }}
                >
                  <button
                    type="button"
                    onClick={() => removePhoto(item.id, i)}
                    className="absolute top-1 right-1 w-5 h-5 bg-black/60 text-white rounded-full flex items-center justify-center active:bg-black"
                  >
                    <X size={11} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* 메모 */}
        <div className="rounded-xl bg-white p-4 shadow-sm">
          <div className="text-[11px] font-bold text-stone-500 uppercase tracking-wider mb-2">
            현장 메모 (선택)
          </div>
          <textarea
            value={memo}
            onChange={(e) => setMemo(e.target.value)}
            placeholder="예) 손잡이 무광 처리 일부 흠집 1개"
            rows={3}
            className="w-full px-3 py-2 text-sm border border-stone-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-stone-900/10 focus:border-stone-400 resize-none"
          />
        </div>
      </div>

      {/* 고정 하단 CTA */}
      <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-stone-200 p-4">
        <button
          type="button"
          onClick={handleComplete}
          disabled={photos.length === 0}
          className="w-full py-3.5 bg-emerald-600 active:bg-emerald-700 text-white font-bold rounded-xl flex items-center justify-center gap-2 disabled:bg-stone-300 disabled:cursor-not-allowed"
        >
          <Check size={18} /> 매칭 완료
        </button>
      </div>
    </div>
  );
}

function ItemPicker({ api }: { api: DemoPOApi }) {
  const { po, select } = api;
  return (
    <div className="min-h-full bg-stone-50">
      <div className="sticky top-0 z-10 bg-white border-b border-stone-200 px-4 py-3">
        <div className="text-[10px] text-stone-400 font-mono">{po.po_number}</div>
        <div className="text-sm font-bold text-stone-900">입고매칭 — 항목 선택</div>
        <div className="text-[11px] text-stone-500 mt-0.5">
          ※ 실제로는 라벨 QR 스캔으로 직진입. 데모에선 리스트에서 선택하세요.
        </div>
      </div>
      <div className="p-3 space-y-2">
        {po.items.length === 0 || !po.items[0].product_name ? (
          <div className="text-center text-stone-400 text-xs py-12">
            STEP 1에서 품목을 먼저 입력하세요.
          </div>
        ) : (
          po.items.map((it) => {
            const seq = it.sticker_no.split('-').pop();
            return (
              <button
                key={it.id}
                type="button"
                onClick={() => select(it.id)}
                className="w-full flex items-center gap-3 p-3 bg-white rounded-xl shadow-sm active:bg-stone-100 text-left"
              >
                <span className="inline-flex items-center justify-center w-10 h-10 rounded-full bg-stone-900 text-white text-sm font-bold">
                  {seq}
                </span>
                <div className="flex-1 min-w-0">
                  <div className="text-sm font-bold text-stone-900 truncate">
                    {it.product_name || '(품목명 없음)'}
                  </div>
                  <div className="text-[11px] text-stone-500 mt-0.5">
                    ¥{it.unit_price} × {it.quantity}
                  </div>
                </div>
                <span
                  className={`px-2 py-0.5 text-[10px] rounded-full border font-medium whitespace-nowrap ${STATUS_TONE[it.inspection_status]}`}
                >
                  {STATUS_LABEL[it.inspection_status]}
                </span>
              </button>
            );
          })
        )}
      </div>
    </div>
  );
}
