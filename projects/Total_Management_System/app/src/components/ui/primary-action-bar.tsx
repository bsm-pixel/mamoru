'use client';

import type { ReactNode } from 'react';

interface Props {
  /** 주 액션 버튼들 */
  children: ReactNode;
  /** 우측 보조 슬롯(선택) — 준비표 등 */
  aside?: ReactNode;
}

/**
 * 상세 패널 상단 고정(sticky) 「다음 할 일」 주 액션 바.
 * - 스크롤해도 상단에 붙어 있어 어느 위치에서든 즉시 처리 (스크롤 피로 제거)
 * - 렌더 여부는 호출부가 hasPrimary 로 판단(빈 fragment도 truthy라 여기서 못 거름)
 * - 파괴적/부차 액션은 넣지 않는다(오클릭 방지) — 하단 유지
 */
export function PrimaryActionBar({ children, aside }: Props) {
  return (
    <div className="sticky top-0 z-20 -mx-0.5 px-0.5 pt-0.5 pb-2 bg-gradient-to-b from-white via-white to-transparent">
      <div className="rounded-xl border border-neutral-900/10 bg-white shadow-sm ring-1 ring-black/5 p-2">
        <div className="flex items-center justify-between mb-1.5 px-0.5">
          <p className="text-[10px] font-bold text-neutral-400 tracking-wide">⚡ 다음 할 일</p>
          {aside}
        </div>
        <div className="flex flex-col gap-1.5">{children}</div>
      </div>
    </div>
  );
}
