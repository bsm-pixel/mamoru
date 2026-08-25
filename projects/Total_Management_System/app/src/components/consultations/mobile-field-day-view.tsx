'use client';

import { useMemo, useState } from 'react';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations } from '@/hooks/use-consultations';
import { formatPhone } from '@/lib/utils/format';
import { Navigation, Copy, ChevronLeft, ChevronRight, Check } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

/** 날짜 포맷: YYYY-MM-DD (KST 로컬 — UTC 슬라이스 시 자정 전후 off-by-one 방지) */
function toDateStr(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function addDays(d: Date, n: number): Date {
  const r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

/** 날짜 라벨: "2/17 (월)" */
function dayLabel(dateStr: string): string {
  const d = new Date(dateStr + 'T00:00:00');
  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  return `${d.getMonth() + 1}/${d.getDate()} (${dayNames[d.getDay()]})`;
}

interface MobileFieldDayViewProps {
  onSelect?: (id: string) => void;
}

export function MobileFieldDayView({ onSelect }: MobileFieldDayViewProps = {}) {
  const today = toDateStr(new Date());
  const [selectedDate, setSelectedDate] = useState(today);
  const [copiedId, setCopiedId] = useState<string | null>(null);

  // 확정된 출장건만 조회
  const { data, isLoading } = useConsultations({
    status: 'confirmed',
    type: 'field_request',
    limit: 100,
  });

  // 날짜 범위 생성 (오늘 기준 ±7일)
  const dateRange = useMemo(() => {
    const dates: string[] = [];
    const base = new Date();
    for (let i = -3; i <= 7; i++) {
      dates.push(toDateStr(addDays(base, i)));
    }
    return dates;
  }, []);

  // 선택 날짜에 해당하는 건 필터 + 시간순 정렬
  const dayConsultations = useMemo(() => {
    if (!data?.consultations) return [];
    return data.consultations
      .filter((c) => c.visit_date === selectedDate)
      .sort((a, b) => (a.visit_time || '99:99').localeCompare(b.visit_time || '99:99'));
  }, [data, selectedDate]);

  const handleNavigation = (c: Consultation) => {
    if (c.latitude && c.longitude) {
      window.open(`kakaomap://route?ep=${c.latitude},${c.longitude}&by=CAR`, '_blank');
    }
  };

  const handleCopyAddress = async (c: Consultation) => {
    const addr = [c.address_road, c.address_detail].filter(Boolean).join(' ');
    if (addr) {
      await navigator.clipboard.writeText(addr);
      setCopiedId(c.id);
      setTimeout(() => setCopiedId(null), 2000);
    }
  };

  return (
    <div className="space-y-4">
      {/* 날짜 스크롤 선택 */}
      <div className="flex gap-1 overflow-x-auto pb-1 -mx-1 px-1">
        {dateRange.map((d) => (
          <button
            key={d}
            onClick={() => setSelectedDate(d)}
            className={`px-3 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition shrink-0 ${
              d === selectedDate
                ? 'bg-terracotta text-cream'
                : d === today
                  ? 'bg-terracotta/10 text-terracotta'
                  : 'bg-card-white text-neutral-500 hover:bg-warm-ivory'
            }`}
          >
            {dayLabel(d)}
          </button>
        ))}
      </div>

      {/* 건수 */}
      <p className="text-xs text-neutral-500">
        {dayLabel(selectedDate)} 출장: <strong>{dayConsultations.length}건</strong>
      </p>

      {/* 목록 */}
      {isLoading ? (
        <div className="space-y-3">
          {Array.from({ length: 3 }).map((_, i) => (
            <Skeleton key={i} className="h-24 w-full" />
          ))}
        </div>
      ) : dayConsultations.length === 0 ? (
        <Card>
          <div className="flex items-center justify-center h-24 text-sm text-neutral-400">
            해당 날짜에 출장 일정이 없습니다
          </div>
        </Card>
      ) : (
        <div className="space-y-3 lg:grid lg:grid-cols-2 lg:gap-4 lg:space-y-0">
          {dayConsultations.map((c) => (
            <Card key={c.id}>
              <div
                className={onSelect ? 'cursor-pointer' : ''}
                onClick={() => onSelect?.(c.id)}
              >
                <div className="flex items-start justify-between">
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-bold text-indigo-black">{c.name}</span>
                      {c.visit_time && (
                        <Badge variant="info">{c.visit_time}</Badge>
                      )}
                    </div>
                    <p className="text-xs text-neutral-500 mt-1">
                      <a href={`tel:${c.phone}`} onClick={(e) => e.stopPropagation()} className="hover:text-blue-600">
                        {formatPhone(c.phone)}
                      </a>
                    </p>
                    <p className="text-xs text-neutral-500 mt-0.5 truncate">
                      {c.address_road} {c.address_detail}
                    </p>
                  </div>
                </div>
              </div>
              <div className="flex gap-2 mt-3">
                <Button
                  variant="primary"
                  size="sm"
                  className="flex-1"
                  onClick={() => handleNavigation(c)}
                  disabled={!c.latitude || !c.longitude}
                >
                  <Navigation size={14} />
                  카카오네비
                </Button>
                <Button
                  variant="secondary"
                  size="sm"
                  className="flex-1"
                  onClick={() => handleCopyAddress(c)}
                >
                  {copiedId === c.id ? <Check size={14} /> : <Copy size={14} />}
                  {copiedId === c.id ? '복사됨' : '주소복사'}
                </Button>
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}
