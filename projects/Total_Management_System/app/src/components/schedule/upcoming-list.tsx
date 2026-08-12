'use client';

/**
 * UpcomingList — 일정 페이지 '다가오는 일정' 리스트 (2026-08-12)
 *   달력 아래 빈 공간을 채움. 다음 8건을 날짜순으로, 각 건에 전화·길찾기(출장) 액션.
 *   /schedule 전용. 공용 useScheduleEvents 재사용.
 */

import { useRouter } from 'next/navigation';
import { useScheduleEvents, flattenEvents, eventCategory, eventName, eventPhone, eventTimeStr, eventHref, eventAddress } from '@/lib/schedule/use-schedule-events';
import { SCHEDULE_COLORS } from '@/lib/schedule/colors';
import { ymd, parseYmd } from '@/lib/schedule/calendar';
import { formatPhone } from '@/lib/utils/format';
import { Phone, Navigation } from 'lucide-react';

export function UpcomingList() {
  const router = useRouter();
  const now = new Date();
  const todayStr = ymd(now.getFullYear(), now.getMonth(), now.getDate());
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 45);
  const endStr = ymd(end.getFullYear(), end.getMonth(), end.getDate());
  const nowHHMM = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

  const { dateMap } = useScheduleEvents(todayStr, endStr);
  const flat = flattenEvents(dateMap);
  const upcoming = flat
    .filter((x) => x.date > todayStr || (x.date === todayStr && (eventTimeStr(x.ev) ?? '99:99') >= nowHHMM))
    .slice(0, 8);

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4">
      <div className="flex items-center justify-between mb-2">
        <h3 className="text-sm font-bold text-stone-900">다가오는 일정</h3>
        <span className="text-[11px] text-stone-400">{upcoming.length}건</span>
      </div>

      {upcoming.length === 0 ? (
        <div className="py-6 text-center text-xs text-stone-400">예정된 일정이 없습니다</div>
      ) : (
        <div className="divide-y divide-stone-100">
          {upcoming.map(({ date, ev }, i) => {
            const cat = eventCategory(ev);
            const color = SCHEDULE_COLORS[cat];
            const p = parseYmd(date);
            const time = eventTimeStr(ev);
            const phone = eventPhone(ev);
            const addr = eventAddress(ev);
            return (
              <div
                key={`${date}-${i}`}
                onClick={() => router.push(eventHref(ev))}
                className="flex items-center gap-2.5 py-2.5 cursor-pointer hover:bg-stone-50 -mx-2 px-2 rounded-lg transition"
              >
                <div className="w-10 shrink-0 text-center">
                  <div className="text-[11px] font-semibold text-stone-600 leading-none">{p.m}/{p.d}</div>
                  <div className="text-[10px] text-stone-400 mt-0.5">({p.dow})</div>
                </div>
                <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${color.dot}`} />
                <span className="text-xs font-semibold text-stone-500 w-10 shrink-0">{time || '-'}</span>
                <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold shrink-0 ${color.badge}`}>{color.label}</span>
                <span className="flex-1 min-w-0 text-sm font-medium text-stone-800 truncate">{eventName(ev)}</span>
                <span className="hidden sm:inline text-[10px] text-stone-400 shrink-0">{phone ? formatPhone(phone) : ''}</span>
                <div className="flex items-center gap-1 shrink-0">
                  {addr && (
                    <a
                      href={`https://map.kakao.com/?q=${encodeURIComponent(addr)}`}
                      target="_blank"
                      rel="noreferrer"
                      onClick={(e) => e.stopPropagation()}
                      className="w-7 h-7 rounded-lg hover:bg-stone-200/70 flex items-center justify-center transition"
                      title="길찾기"
                      aria-label="길찾기"
                    >
                      <Navigation size={14} className="text-stone-500" />
                    </a>
                  )}
                  {phone && (
                    <a
                      href={`tel:${phone}`}
                      onClick={(e) => e.stopPropagation()}
                      className="w-7 h-7 rounded-lg hover:bg-stone-200/70 flex items-center justify-center transition"
                      title="전화"
                      aria-label="전화"
                    >
                      <Phone size={14} className="text-stone-500" />
                    </a>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
