'use client';

/**
 * /sales 페이지 디자인 프리뷰 — 사장님 비교 선택용 (운영 X)
 *
 * 2026-05-26 Phase G-3: 매출 카드 3안 + 목록 카드 4안 동시 렌더. 사장님이 한 화면에서 비교 → 선택 통보 → G-4 적용.
 * 실제 데이터 호출 X — 모두 하드코딩 임시 데이터.
 */

import { Topbar } from '@/components/layout/topbar';
import { Calendar, TrendingUp, AlertCircle } from 'lucide-react';
import { formatKRW } from '@/lib/utils/format';

// ─── 임시 데이터 ───────────────────────────────────────────────
const STATS = {
  customer: {
    week: { amount: 1190000, count: 4 },
    month: { amount: 6567010, count: 27 },
    outstanding: 280000,
  },
  partner: {
    week: { amount: 2400000, count: 3 },
    month: { amount: 9800000, count: 15 },
    outstanding: 1500000,
  },
};

type SampleRow = {
  id: string;
  sourceType: 'sale' | 'delivery';
  name: string;
  date: string; // 5월 26일
  amount: number;
  paymentStatus: 'paid' | 'unpaid' | 'partial';
  paymentMethod: string; // 카드 / 이체 / 현금
  channel?: 'offline' | 'talk'; // B2C 만
  partnerType?: 'dealer' | 'academy'; // B2B 만
  shipped?: boolean;
  delivered?: boolean;
  cancelled?: boolean;
};

const SAMPLE_ROWS: SampleRow[] = [
  { id: '1', sourceType: 'sale', name: '서윤 점장님(박미애)', date: '5월 26일', amount: 100000, paymentStatus: 'paid', paymentMethod: '카드', channel: 'talk', delivered: true },
  { id: '2', sourceType: 'sale', name: '전아연', date: '5월 25일', amount: 780000, paymentStatus: 'paid', paymentMethod: '이체', channel: 'talk', shipped: true },
  { id: '3', sourceType: 'sale', name: '김승한', date: '5월 18일', amount: 280000, paymentStatus: 'unpaid', paymentMethod: '카드', channel: 'offline' },
  { id: '4', sourceType: 'sale', name: '곽경진', date: '5월 25일', amount: 60000, paymentStatus: 'partial', paymentMethod: '이체', channel: 'offline' },
  { id: '5', sourceType: 'delivery', name: '○○매장', date: '5월 24일', amount: 1240000, paymentStatus: 'paid', paymentMethod: '이체', partnerType: 'dealer' },
  { id: '6', sourceType: 'delivery', name: '△△아카데미', date: '5월 23일', amount: 580000, paymentStatus: 'unpaid', paymentMethod: '이체', partnerType: 'academy' },
  { id: '7', sourceType: 'sale', name: '취소건 예시', date: '5월 20일', amount: 150000, paymentStatus: 'paid', paymentMethod: '카드', channel: 'offline', cancelled: true },
];

// ─── 페이지 ────────────────────────────────────────────────────
export default function SalesDesignPreviewPage() {
  return (
    <>
      <Topbar title="디자인 프리뷰 (사장님 선택용)" />
      <div className="px-6 py-6 space-y-10 max-w-[1200px] mx-auto">
        {/* 안내 */}
        <div className="px-4 py-3 rounded-lg bg-yellow-50 border border-yellow-200 text-sm text-yellow-800">
          💡 <strong>디자인 비교 프리뷰</strong> — 운영 데이터 X, UI 미리보기만.
          매출 카드 1안 + 목록 카드 1안 선택해서 클로드에게 알려주세요. (예: "매출은 안 2, 목록은 안 C")
        </div>

        {/* § A 매출 카드 옵션 */}
        <section>
          <h2 className="text-base font-bold text-indigo-black mb-3">§ A. 상단 매출 카드 (3안)</h2>

          <SubLabel label="안 1 — Stripe/Linear 스타일 (큰 숫자 중심)" />
          <Option1Stats />

          <SubLabel label="안 2 — 좌측 accent line (HubCard 스타일)" />
          <Option2Stats />

          <SubLabel label="안 3 — 어두운 카드 (복원수리 스타일)" />
          <Option3Stats />
        </section>

        {/* § B 목록 카드 옵션 */}
        <section>
          <h2 className="text-base font-bold text-indigo-black mb-3">§ B. 좌측 목록 카드 (4안)</h2>

          <SubLabel label="안 A — 도트 + 색 줄 (가장 미니멈)" />
          <Wrap>{SAMPLE_ROWS.map((r) => <RowA key={r.id} row={r} />)}</Wrap>

          <SubLabel label="안 B — 우측 위계 정돈" />
          <Wrap>{SAMPLE_ROWS.map((r) => <RowB key={r.id} row={r} />)}</Wrap>

          <SubLabel label="안 C — 통합 1개 뱃지 (Linear 스타일)" />
          <Wrap>{SAMPLE_ROWS.map((r) => <RowC key={r.id} row={r} />)}</Wrap>

          <SubLabel label="안 D — 미니 그리드 (Notion 스타일)" />
          <Wrap>{SAMPLE_ROWS.map((r) => <RowD key={r.id} row={r} />)}</Wrap>
        </section>
      </div>
    </>
  );
}

// ─── 공용 helpers ───────────────────────────────────────────────
function SubLabel({ label }: { label: string }) {
  return (
    <div className="flex items-center gap-2 mt-6 mb-2">
      <span className="px-2.5 py-1 rounded-md bg-neutral-900 text-white text-xs font-semibold">{label}</span>
    </div>
  );
}
function Wrap({ children }: { children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-xl border border-neutral-200 overflow-hidden">
      <div className="divide-y divide-neutral-100">{children}</div>
    </div>
  );
}
function statusText(r: SampleRow): string {
  if (r.cancelled) return '취소';
  if (r.paymentStatus === 'unpaid') return '미결제';
  if (r.paymentStatus === 'partial') return '부분결제';
  return '결제완료';
}
function channelText(r: SampleRow): string {
  if (r.sourceType === 'delivery') {
    return r.partnerType === 'dealer' ? '딜러' : r.partnerType === 'academy' ? '아카데미' : '거래처';
  }
  return r.channel === 'talk' ? '온라인상담' : '오프라인';
}
function operationText(r: SampleRow): string | null {
  if (r.cancelled) return null;
  if (r.delivered) return '판매완료';
  if (r.shipped) return '배송중';
  return null;
}

// ═══════════════════════════════════════════════════════════════
// § A 매출 카드 옵션
// ═══════════════════════════════════════════════════════════════

/** 안 1 — Stripe/Linear 스타일 (큰 숫자 중심) */
function Option1Stats() {
  return (
    <div className="grid grid-cols-2 gap-4">
      {([
        { title: '고객 (B2C)', data: STATS.customer },
        { title: '거래처 (B2B)', data: STATS.partner },
      ] as const).map((sec) => (
        <div key={sec.title} className="bg-white rounded-xl border border-neutral-200 p-5">
          <div className="text-[11px] font-semibold text-neutral-500 uppercase tracking-wider mb-2">{sec.title}</div>
          <div className="text-3xl font-bold text-neutral-900 tracking-tight">{formatKRW(sec.data.month.amount)}</div>
          <div className="text-xs text-neutral-500 mt-1">이번달 · {sec.data.month.count}건</div>
          <div className="mt-4 pt-3 border-t border-neutral-100 flex items-center justify-between text-xs">
            <span className="text-neutral-500">이번주 <span className="font-semibold text-neutral-700">{formatKRW(sec.data.week.amount)}</span></span>
            <span className="text-red-500">미수금 <span className="font-semibold">{formatKRW(sec.data.outstanding)}</span> ⚠</span>
          </div>
        </div>
      ))}
    </div>
  );
}

/** 안 2 — 좌측 accent line */
function Option2Stats() {
  return (
    <div className="grid grid-cols-2 gap-4">
      {([
        { title: '고객 (B2C)', data: STATS.customer },
        { title: '거래처 (B2B)', data: STATS.partner },
      ] as const).map((sec) => (
        <div key={sec.title} className="bg-white rounded-lg border border-neutral-200 overflow-hidden flex">
          <div className="w-1 bg-neutral-900" />
          <div className="flex-1 p-4">
            <div className="text-sm font-bold text-neutral-900 mb-3">{sec.title}</div>
            <div className="space-y-1.5 text-sm">
              <div className="flex items-center justify-between"><span className="text-neutral-500">이번달</span><span><span className="font-bold">{formatKRW(sec.data.month.amount)}</span><span className="text-xs text-neutral-400 ml-2">{sec.data.month.count}건</span></span></div>
              <div className="flex items-center justify-between"><span className="text-neutral-500">이번주</span><span><span className="font-semibold text-neutral-700">{formatKRW(sec.data.week.amount)}</span><span className="text-xs text-neutral-400 ml-2">{sec.data.week.count}건</span></span></div>
              <div className="flex items-center justify-between pt-1.5 border-t border-neutral-100"><span className="text-red-500">미수금</span><span className="font-semibold text-red-600">{formatKRW(sec.data.outstanding)} ⚠</span></div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

/** 안 3 — 어두운 카드 (복원수리 스타일) */
function Option3Stats() {
  return (
    <div className="grid grid-cols-2 gap-4">
      {([
        { title: '고객 (B2C)', data: STATS.customer },
        { title: '거래처 (B2B)', data: STATS.partner },
      ] as const).map((sec) => (
        <div key={sec.title} className="bg-neutral-900 text-white rounded-xl p-5">
          <div className="text-[11px] font-semibold uppercase tracking-wider opacity-60 mb-2">{sec.title}</div>
          <div className="text-3xl font-bold tracking-tight">{formatKRW(sec.data.month.amount)}</div>
          <div className="text-xs opacity-70 mt-1">이번달 · {sec.data.month.count}건</div>
          <div className="mt-4 pt-3 border-t border-white/10 grid grid-cols-2 gap-2 text-xs">
            <div><span className="opacity-60">이번주</span><div className="font-semibold mt-0.5">{formatKRW(sec.data.week.amount)}</div></div>
            <div><span className="text-amber-300">미수금</span><div className="font-semibold mt-0.5 text-amber-300">{formatKRW(sec.data.outstanding)}</div></div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// § B 목록 카드 옵션
// ═══════════════════════════════════════════════════════════════

/** 안 A — 도트 + 색 줄 (가장 미니멈) */
function RowA({ row }: { row: SampleRow }) {
  const stripColor = row.cancelled ? 'bg-neutral-300' :
    row.paymentStatus === 'unpaid' ? 'bg-red-400' :
    row.paymentStatus === 'partial' ? 'bg-yellow-400' : 'bg-transparent';
  return (
    <div className={`flex items-center gap-4 px-4 py-3 ${row.cancelled ? 'opacity-50' : ''}`}>
      <div className={`w-1 h-10 rounded ${stripColor}`} />
      <div className="flex-1 min-w-0">
        <div className={`text-sm font-semibold text-indigo-black truncate ${row.cancelled ? 'line-through' : ''}`}>{row.name}</div>
        <div className="text-xs text-neutral-500 mt-0.5">{row.date} · {row.paymentMethod} · {channelText(row)}</div>
      </div>
      {/* 우측 운영 상태 도트 */}
      <div className="flex items-center gap-2 shrink-0">
        {row.shipped && <span className="w-2 h-2 rounded-full bg-green-500" title="배송중" />}
        {row.delivered && <span className="w-2 h-2 rounded-full bg-neutral-400" title="판매완료" />}
        <span className={`text-sm font-bold ${row.cancelled ? 'line-through text-neutral-400' : ''}`}>{formatKRW(row.amount)}</span>
      </div>
    </div>
  );
}

/** 안 B — 우측 위계 정돈 */
function RowB({ row }: { row: SampleRow }) {
  return (
    <div className={`flex items-start gap-4 px-4 py-3 ${row.cancelled ? 'opacity-50' : ''}`}>
      <div className="flex-1 min-w-0">
        <div className={`text-sm font-semibold text-indigo-black truncate ${row.cancelled ? 'line-through' : ''}`}>{row.name}</div>
        <div className="text-xs text-neutral-500 mt-0.5">{row.date} · {row.paymentMethod}</div>
      </div>
      <div className="text-right shrink-0">
        <div className={`text-sm font-bold ${row.cancelled ? 'line-through text-neutral-400' : ''}`}>{formatKRW(row.amount)}</div>
        <div className="text-[11px] text-neutral-500 mt-0.5">
          {row.cancelled ? '취소' : statusText(row)} · {channelText(row)}
          {operationText(row) && <> · <span className={row.shipped ? 'text-green-600' : ''}>{operationText(row)}</span></>}
        </div>
      </div>
    </div>
  );
}

/** 안 C — 통합 1개 뱃지 (Linear 스타일) */
function RowC({ row }: { row: SampleRow }) {
  const dotColor = row.cancelled ? 'bg-neutral-300' :
    row.paymentStatus === 'unpaid' ? 'bg-red-500' :
    row.paymentStatus === 'partial' ? 'bg-yellow-500' :
    row.shipped ? 'bg-green-500' : 'bg-neutral-400';
  const dotFilled = !row.cancelled && row.paymentStatus === 'paid';
  return (
    <div className={`flex items-center gap-4 px-4 py-3 ${row.cancelled ? 'opacity-50' : ''}`}>
      <div className="flex-1 min-w-0">
        <div className={`text-sm font-semibold text-indigo-black truncate ${row.cancelled ? 'line-through' : ''}`}>{row.name}</div>
        <div className="flex items-center gap-1.5 text-xs text-neutral-500 mt-0.5">
          <span className={`inline-block w-2 h-2 rounded-full ${dotFilled ? dotColor : `border-2 border-current bg-transparent ${dotColor.replace('bg-', 'text-')}`}`} />
          <span>{statusText(row)}</span>
          <span className="text-neutral-300">·</span>
          <span>{channelText(row)}</span>
          {operationText(row) && (<><span className="text-neutral-300">·</span><span>{operationText(row)}</span></>)}
          <span className="text-neutral-300">·</span>
          <span className="text-neutral-400">{row.date.replace(/^\d+월 /, '').replace('일', '')}/{row.date.match(/^(\d+)월/)?.[1]}</span>
        </div>
      </div>
      <div className={`text-sm font-bold shrink-0 ${row.cancelled ? 'line-through text-neutral-400' : ''}`}>{formatKRW(row.amount)}</div>
    </div>
  );
}

/** 안 D — 미니 그리드 (Notion 스타일) */
function RowD({ row }: { row: SampleRow }) {
  const tagColor = row.sourceType === 'delivery'
    ? 'bg-indigo-50 text-indigo-700 border-indigo-200'
    : row.channel === 'talk'
      ? 'bg-yellow-50 text-yellow-700 border-yellow-200'
      : 'bg-neutral-50 text-neutral-600 border-neutral-200';
  return (
    <div className={`flex items-center gap-4 px-4 py-3 ${row.cancelled ? 'opacity-50' : ''}`}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <span className={`text-sm font-semibold text-indigo-black truncate ${row.cancelled ? 'line-through' : ''}`}>{row.name}</span>
          <span className={`text-[10px] font-medium px-1.5 py-0.5 rounded border ${tagColor}`}>{channelText(row)}</span>
        </div>
        <div className="text-xs text-neutral-500 mt-0.5">{row.date} · {row.paymentMethod} · {row.cancelled ? '취소' : statusText(row)}{operationText(row) && ` · ${operationText(row)}`}</div>
      </div>
      <div className={`text-sm font-bold shrink-0 ${row.cancelled ? 'line-through text-neutral-400' : ''}`}>{formatKRW(row.amount)}</div>
    </div>
  );
}
