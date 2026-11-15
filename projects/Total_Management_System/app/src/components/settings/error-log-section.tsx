'use client';

import { useQuery } from '@tanstack/react-query';
import { RefreshCw } from 'lucide-react';
import { formatDateTimeShort } from '@/lib/utils/format';
import { Fold } from './fold';

/**
 * 오류 모아 보기 (2026-11-15) — 설정 > 시스템
 * 화면 오류 · 알림톡 발송 실패 · 동기화 실패를 한곳에서 본다(최근 50건). 기록은 lib/log/error-log.ts 가 남긴다.
 */
const KIND_LABEL: Record<string, string> = {
  error_screen: '화면',
  error_alimtalk: '알림톡 발송',
  error_api: '서버',
  imweb_orders: '아임웹 주문 동기화',
  imweb_products: '아임웹 상품 동기화',
};

interface Row { id: string; sync_type: string; status: string; error_message: string | null; started_at: string }

export function ErrorLogSection() {
  const { data, isFetching, refetch } = useQuery({
    queryKey: ['error-log'],
    staleTime: 30_000,
    queryFn: async () => {
      const res = await fetch('/api/log/error');
      if (!res.ok) return { rows: [] as Row[], recent: 0 };
      const j = await res.json();
      const rows = (j.errors || []) as Row[];
      // 최근 7일 건수는 받아 올 때 한 번만 센다(그릴 때마다 지금 시각을 읽지 않게)
      const since = Date.now() - 7 * 86400000;
      return { rows, recent: rows.filter((r) => new Date(r.started_at).getTime() >= since).length };
    },
  });
  const rows = data?.rows || [];
  const recent = data?.recent || 0;

  return (
    <Fold title={`오류 모아 보기${recent ? ` — 최근 7일 ${recent}건` : ''}`} desc="화면 오류 · 알림톡 발송 실패 · 동기화 실패 (최근 50건). 고객 이름 · 번호는 남기지 않습니다">
      <div className="flex justify-end">
        <button type="button" onClick={() => refetch()} className="h-8 px-2.5 rounded-lg border border-neutral-200 text-xs font-semibold text-neutral-600 hover:bg-neutral-50 inline-flex items-center gap-1">
          <RefreshCw size={12} className={isFetching ? 'animate-spin' : ''} />새로 고침
        </button>
      </div>
      {rows.length === 0 ? (
        <p className="py-4 text-center text-xs text-neutral-400">남은 오류가 없습니다</p>
      ) : (
        <ul className="divide-y divide-neutral-100">
          {rows.map((r) => (
            <li key={r.id} className="py-2 flex items-start gap-3">
              <span className="shrink-0 w-24 text-[11px] text-neutral-400 tabular-nums pt-0.5">{formatDateTimeShort(r.started_at)}</span>
              <span className="shrink-0 px-1.5 py-0.5 rounded text-[10.5px] font-bold bg-error-soft text-error">{KIND_LABEL[r.sync_type] || r.sync_type}</span>
              <span className="min-w-0 text-xs text-neutral-700 break-words">{r.error_message || '내용 없음'}</span>
            </li>
          ))}
        </ul>
      )}
    </Fold>
  );
}
