'use client';

import type { ReactNode } from 'react';

/**
 * 모바일 폭(390px) 프레임 — 디자인 모니터 안에서 모바일 화면을 PC와 나란히 미리보기.
 * 실제 운영에선 사용 X (반응형 페이지가 그대로 폰에 뜨면 됨).
 */
export function MobileFrame({
  children,
  title,
}: {
  children: ReactNode;
  title?: string;
}) {
  return (
    <div className="inline-flex flex-col items-center">
      {title && (
        <div className="text-[11px] text-stone-500 mb-2 font-medium">{title}</div>
      )}
      <div
        className="relative bg-stone-950 rounded-[36px] p-2 shadow-2xl"
        style={{ width: 406 }}
      >
        <div className="absolute top-2 left-1/2 -translate-x-1/2 w-24 h-5 bg-stone-950 rounded-b-2xl z-20" />
        <div
          className="relative bg-stone-50 rounded-[28px] overflow-hidden"
          style={{ width: 390, height: 780 }}
        >
          <div className="absolute inset-0 overflow-y-auto">{children}</div>
        </div>
      </div>
      <div className="text-[10px] text-stone-400 mt-2">390 × 780</div>
    </div>
  );
}
