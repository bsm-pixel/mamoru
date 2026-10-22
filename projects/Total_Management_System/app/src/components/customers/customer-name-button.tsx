'use client';

import { useState } from 'react';
import { CustomerQuickModal } from '@/components/customers/customer-quick-modal';

/**
 * 고객 이름 → 「고객 정보」 모달(메모·상담 메모·최근 거래) (2026-10-22)
 *
 * 상담관리 목록·상세처럼 **카드 전체가 클릭(상세 열기)** 인 곳에서 이름만 따로 눌러 고객을 미리 본다.
 *   · 이름 클릭은 카드 클릭으로 번지지 않는다(stopPropagation)
 *   · 모달도 카드 안에 그려지므로 모달 안 클릭이 카드로 번지지 않게 한 겹 감싼다
 *   · customerId 가 없으면 그냥 글자(누를 수 없음)
 */
export function CustomerNameButton({ customerId, className = '', children }: {
  customerId?: string | null;
  className?: string;
  children: React.ReactNode;
}) {
  const [open, setOpen] = useState(false);
  if (!customerId) return <span className={className}>{children}</span>;
  return (
    <>
      <button
        type="button"
        title="고객 정보·메모 보기"
        onClick={(e) => { e.stopPropagation(); setOpen(true); }}
        className={`${className} text-left underline decoration-dotted decoration-neutral-300 underline-offset-4 hover:decoration-neutral-700 transition`}
      >
        {children}
      </button>
      {open && (
        <span className="contents" onClick={(e) => e.stopPropagation()} onKeyDown={(e) => e.stopPropagation()}>
          <CustomerQuickModal customerId={customerId} open onClose={() => setOpen(false)} />
        </span>
      )}
    </>
  );
}
