'use client';

/**
 * AllConsultationsList — 매장방문 + 출장요청 + 온라인상담 통합 리스트
 *
 * 시안 B 톤 (2026-05-27 사장님 채택):
 *   - 좌측 색 줄 (매장 emerald / 출장 violet / 톡 blue)
 *   - 좌측 타입 아이콘 박스
 *   - 시간순 정렬 (visit_date_asc), 활성 6상태
 *   - 검색 + 페이지네이션
 *
 * 사장님 룰: 전체 탭은 모든 활성 상담을 한눈에 보기 위한 진입점 (기본값).
 */

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations } from '@/hooks/use-consultations';
import { activityDisplay } from '@/lib/customer/display';
import {
  formatPhone,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { Search, ChevronLeft, ChevronRight, Store, Truck, MessageCircle, ArrowRight, MapPin } from 'lucide-react';

const ACTIVE_STATUSES = [
  'pending_admin',
  'suggested',
  'confirmed',
  'in_progress',
  'reschedule_requested',
  'change_requested',
];

const TYPE_CONFIG = {
  store_visit:  { label: '매장', icon: Store,         color: 'emerald', bar: 'bg-emerald-500', bg: 'bg-emerald-50', text: 'text-emerald-700', iconColor: 'text-emerald-600' },
  field_request:{ label: '출장', icon: Truck,         color: 'violet',  bar: 'bg-violet-500',  bg: 'bg-violet-50',  text: 'text-violet-700',  iconColor: 'text-violet-600' },
  talk_consult: { label: '톡상담', icon: MessageCircle,color: 'blue',    bar: 'bg-blue-500',    bg: 'bg-blue-50',    text: 'text-blue-700',    iconColor: 'text-blue-600' },
} as const;

export function AllConsultationsList({ onSelect }: { onSelect?: (id: string) => void } = {}) {
  const router = useRouter();
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);

  const { data, isLoading } = useConsultations({
    statuses: ACTIVE_STATUSES,
    search,
    page,
    limit: 30,
    orderBy: 'visit_date_asc',
  });

  const consultations = data?.consultations || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 30);

  return (
    <div className="space-y-3">
      {/* 검색 */}
      <div className="relative">
        <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-400" />
        <input
          type="text"
          value={search}
          onChange={(e) => { setSearch(e.target.value); setPage(1); }}
          placeholder="이름, 전화번호 검색"
          className="w-full h-9 pl-9 pr-3 rounded-xl border border-stone-200 bg-white text-sm text-stone-800 placeholder:text-stone-400 focus:outline-none focus:border-stone-400 transition"
        />
        <span className="absolute right-3 top-1/2 -translate-y-1/2 text-xs text-stone-400">{total}건</span>
      </div>

      {/* 목록 */}
      {isLoading ? (
        <div className="space-y-2">
          {Array.from({ length: 6 }).map((_, i) => (
            <Skeleton key={i} className="h-16 w-full rounded-2xl" />
          ))}
        </div>
      ) : consultations.length === 0 ? (
        <Card>
          <div className="flex items-center justify-center h-32 text-sm text-stone-400">
            진행 중인 상담이 없습니다
          </div>
        </Card>
      ) : (
        <div className="space-y-2">
          {consultations.map((c) => {
            const type = c.consultation_type as keyof typeof TYPE_CONFIG;
            const cfg = TYPE_CONFIG[type] || TYPE_CONFIG.talk_consult;
            const Icon = cfg.icon;
            return (
              <div
                key={c.id}
                onClick={() => onSelect ? onSelect(c.id) : router.push(`/consultations/${c.id}`)}
                className="bg-white rounded-2xl border border-stone-200 overflow-hidden hover:border-stone-300 transition flex items-stretch group cursor-pointer"
              >
                {/* 좌측 색 줄 (타입별) */}
                <div className={`w-1 ${cfg.bar}`} />
                <div className="flex-1 p-3 flex items-center gap-3 min-w-0">
                  {/* 좌측 아이콘 박스 */}
                  <div className={`w-10 h-10 rounded-lg ${cfg.bg} flex items-center justify-center shrink-0`}>
                    <Icon size={16} className={cfg.iconColor} />
                  </div>
                  {/* 본문 */}
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-0.5 flex-wrap">
                      <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold shrink-0 ${cfg.bg} ${cfg.text}`}>{cfg.label}</span>
                      <p className="text-sm font-semibold text-stone-800 truncate">{activityDisplay(c.activity_name, c.name)}</p>
                      <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold shrink-0 ${CONSULTATION_STATUS_COLOR[c.status] || 'bg-stone-100 text-stone-600'}`}>
                        {CONSULTATION_STATUS_LABEL[c.status] || c.status}
                      </span>
                    </div>
                    <div className="flex items-center gap-2 text-[11px] text-stone-500 flex-wrap">
                      {c.visit_date && (
                        <span className="font-semibold text-stone-700">
                          {c.visit_date}{c.visit_time ? ` ${c.visit_time}` : ''}
                        </span>
                      )}
                      {c.visit_date && c.phone && <span className="text-stone-300">·</span>}
                      {c.phone && <span>{formatPhone(c.phone)}</span>}
                      {c.address_road && (
                        <>
                          <span className="text-stone-300">·</span>
                          <span className="flex items-center gap-0.5"><MapPin size={10} />{c.address_road}</span>
                        </>
                      )}
                    </div>
                  </div>
                  <ArrowRight size={14} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition shrink-0" />
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* 페이지네이션 */}
      {totalPages > 1 && (
        <div className="flex items-center justify-center gap-2 pt-1">
          <Button variant="ghost" size="sm" disabled={page <= 1} onClick={() => setPage(page - 1)}>
            <ChevronLeft size={16} />
          </Button>
          <span className="text-sm text-stone-500">{page} / {totalPages}</span>
          <Button variant="ghost" size="sm" disabled={page >= totalPages} onClick={() => setPage(page + 1)}>
            <ChevronRight size={16} />
          </Button>
        </div>
      )}
    </div>
  );
}
