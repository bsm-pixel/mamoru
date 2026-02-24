'use client';

import { useState } from 'react';
import Link from 'next/link';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from './repair-status-badge';
import { useRepairs } from '@/hooks/use-repairs';
import { formatKRW, formatPhone, formatRelative } from '@/lib/utils/format';
import { Search, ChevronLeft, ChevronRight, Scissors } from 'lucide-react';

type TabKey = 'all' | 'pickup' | 'inspect' | 'repair' | 'shipping' | 'completed' | 'cancelled';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'all', label: '전체' },
  { key: 'pickup', label: '접수/수거' },
  { key: 'inspect', label: '검수/입금' },
  { key: 'repair', label: '수리중' },
  { key: 'shipping', label: '출고/배송' },
  { key: 'completed', label: '완료' },
  { key: 'cancelled', label: '취소' },
];

export function RepairList() {
  const [activeTab, setActiveTab] = useState<TabKey>('all');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const limit = 20;

  const { data, isLoading } = useRepairs({
    status: activeTab === 'all' ? undefined : activeTab,
    search: search || undefined,
    page,
    limit,
  });

  const totalPages = data ? Math.ceil(data.total / limit) : 0;

  // 경과일 계산
  const getDaysElapsed = (receivedAt: string) => {
    const diff = Date.now() - new Date(receivedAt).getTime();
    return Math.floor(diff / (1000 * 60 * 60 * 24));
  };

  return (
    <div className="space-y-4">
      {/* 검색 */}
      <div className="relative">
        <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
        <input
          type="text"
          value={search}
          onChange={(e) => { setSearch(e.target.value); setPage(1); }}
          placeholder="이름, 전화번호, 접수번호 검색..."
          className="w-full h-10 pl-10 pr-4 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
        />
      </div>

      {/* 탭 */}
      <div className="flex gap-1 overflow-x-auto border-b border-neutral-200 scrollbar-hide">
        {TABS.map((tab) => (
          <button
            key={tab.key}
            onClick={() => { setActiveTab(tab.key); setPage(1); }}
            className={`shrink-0 px-3 py-2 text-sm font-semibold border-b-2 transition ${
              activeTab === tab.key
                ? 'border-terracotta text-terracotta'
                : 'border-transparent text-neutral-500 hover:text-neutral-700'
            }`}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {/* 카운트 */}
      {data && (
        <p className="text-xs text-neutral-500">
          총 {data.total}건
        </p>
      )}

      {/* 목록 */}
      {isLoading ? (
        <div className="space-y-3">
          {[...Array(5)].map((_, i) => (
            <Skeleton key={i} className="h-24 w-full rounded-xl" />
          ))}
        </div>
      ) : !data?.repairs.length ? (
        <div className="flex flex-col items-center justify-center py-16 text-neutral-400">
          <Scissors size={32} className="mb-2 opacity-50" />
          <p className="text-sm">복원수리 건이 없습니다</p>
        </div>
      ) : (
        <div className="space-y-2">
          {data.repairs.map((r) => {
            const days = getDaysElapsed(r.received_at);
            return (
              <Link key={r.id} href={`/repairs/${r.id}`}>
                <Card className="hover:bg-neutral-50 transition cursor-pointer">
                  <div className="flex items-start justify-between gap-3">
                    <div className="flex-1 min-w-0">
                      {/* 접수번호 + 상태 */}
                      <div className="flex items-center gap-2 mb-1">
                        <span className="text-xs font-mono text-neutral-500">{r.as_id}</span>
                        <RepairStatusBadge status={r.status} />
                      </div>
                      {/* 고객 정보 */}
                      <p className="text-sm font-semibold truncate">{r.name}</p>
                      <p className="text-xs text-neutral-500">{formatPhone(r.phone)}</p>
                      {/* 가위 수량 */}
                      <div className="flex items-center gap-3 mt-1.5 text-xs text-neutral-600">
                        {r.qty_mamoru > 0 && (
                          <span>마모루 {r.qty_mamoru}자루</span>
                        )}
                        {r.qty_other > 0 && (
                          <span>타사 {r.qty_other}자루</span>
                        )}
                        {r.total_amount > 0 && (
                          <span className="font-medium text-terracotta-deep">
                            {formatKRW(r.total_amount)}
                          </span>
                        )}
                      </div>
                    </div>
                    {/* 우측: 진행방식 + 경과 */}
                    <div className="text-right shrink-0">
                      {r.proceed_type && (
                        <span className="text-xs text-neutral-500">{r.proceed_type}</span>
                      )}
                      <p className="text-[11px] text-neutral-400 mt-1">
                        {formatRelative(r.received_at)}
                      </p>
                      {days > 0 && r.status !== 'completed' && r.status !== 'cancelled' && (
                        <p className={`text-[11px] mt-0.5 ${days >= 7 ? 'text-error font-medium' : 'text-neutral-400'}`}>
                          {days}일 경과
                        </p>
                      )}
                    </div>
                  </div>
                </Card>
              </Link>
            );
          })}
        </div>
      )}

      {/* 페이지네이션 */}
      {totalPages > 1 && (
        <div className="flex items-center justify-center gap-2 pt-2">
          <Button
            variant="ghost"
            size="sm"
            disabled={page <= 1}
            onClick={() => setPage(page - 1)}
          >
            <ChevronLeft size={14} />
          </Button>
          <span className="text-sm text-neutral-500">
            {page} / {totalPages}
          </span>
          <Button
            variant="ghost"
            size="sm"
            disabled={page >= totalPages}
            onClick={() => setPage(page + 1)}
          >
            <ChevronRight size={14} />
          </Button>
        </div>
      )}
    </div>
  );
}
