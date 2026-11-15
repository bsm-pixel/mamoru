'use client';

import { useEffect } from 'react';
import Link from 'next/link';

/**
 * 한 화면에서 오류가 났을 때 (2026-11-15)
 * 전엔 화면 전체가 하얗게 됐다. 이제 왼쪽 메뉴는 그대로 두고 이 화면만 "문제가 생겼습니다 · 다시 시도"를 보여 준다.
 */
export default function DashboardError({ error, reset }: { error: Error & { digest?: string }; reset: () => void }) {
  useEffect(() => {
    console.error('[화면 오류]', error);
    // 설정 > 시스템 「오류 모아 보기」에 남긴다 (실패해도 무시)
    fetch('/api/log/error', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: error?.message || String(error), where: window.location.pathname }),
    }).catch(() => {});
  }, [error]);
  return (
    <div className="flex-1 flex items-center justify-center p-6 min-h-[60vh]">
      <div className="text-center max-w-sm">
        <p className="text-sm font-bold text-stone-900">이 화면에 문제가 생겼습니다</p>
        <p className="mt-1.5 text-xs text-neutral-500">넣던 내용이 저장되지 않았을 수 있습니다. 다시 시도해도 같으면 화면 이름과 함께 알려 주세요.</p>
        {error?.message && <p className="mt-3 px-3 py-2 rounded-lg bg-neutral-100 text-[11px] text-neutral-500 break-words text-left">{error.message.slice(0, 200)}</p>}
        <div className="mt-5 flex items-center justify-center gap-2">
          <button type="button" onClick={reset} className="h-10 px-5 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800">다시 시도</button>
          <Link href="/dashboard" className="h-10 px-5 rounded-lg border border-neutral-200 text-sm font-semibold text-neutral-700 hover:bg-neutral-50 inline-flex items-center">대시보드로</Link>
        </div>
      </div>
    </div>
  );
}
