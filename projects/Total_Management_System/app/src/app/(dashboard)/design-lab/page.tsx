'use client';

/**
 * /design-lab — TMS 디자인 모니터 (사장님 + 클로드 협업 도구)
 *
 * 2026-05-26 사장님 운영 룰 (필독):
 *   ▶ 디자인 모니터는 그때그때 진행 중인 디자인 작업만 표시
 *   ▶ 작업 완료 (사장님 결정 + 실제 페이지 적용) 후 § 자동 삭제
 *   ▶ 영구 박제 X (실제 적용된 디자인은 운영 페이지 + memory/docs 에 박제)
 *
 * 진행 작업: § 대시보드 리모델 — 시안 B+ (압축 트렌디 버전)
 * 플랜: C:/Users/user/.claude/plans/tms-distributed-squid.md
 *
 * TMS 디자인 방향 (2026-05-26 사장님 결정):
 *   - 마모루 브랜드 가이드 100% 추종 X — 최신 트렌드 + 작업효율 극대화 우선
 *   - 컬러감만 모노크롬 베이스 + 절제된 상태색 (memory/feedback_tms_design_direction.md)
 */

import { Topbar } from '@/components/layout/topbar';
import {
  ShoppingCart, MessageSquare, Wrench, Store,
  CheckCircle2, AlertTriangle,
  PackageX, Truck, PackageOpen, Star, ClipboardList,
  ChevronLeft, ChevronRight, ArrowRight, Plus,
} from 'lucide-react';

/* ──────────────────────────────────────────────────────────
 * Mock 데이터 (정적 — 운영 호출 X)
 * ──────────────────────────────────────────────────────── */
const MOCK = {
  monthGoal: 15_000_000,
  current: 11_200_000,
  b2c: 3_500_000,
  b2b: 4_800_000,
  repair: 2_900_000,
  pct: 75,
  orders: { payDone: 12, preparing: 3, shipping: 2, delivered: 8, monthAmount: 2_400_000 },
  consultations: { newIntake: 5, confirmed: 8, needAction: 2 },
  repairs: { intakeNew: 3, pendingInbound: 2, workingCount: 4, readyToShip: 6, monthBags: 15 },
  sales: { monthCount: 32 },
  todayDate: { month: 5, day: 26, weekday: '화' },
  today: [
    { id: '1', type: 'store', name: '김미용', subtitle: '강남', time: '14:00' },
    { id: '2', type: 'field', name: '박헤어', subtitle: '판교 출장', time: '15:30' },
    { id: '3', type: 'repair', name: '이살롱', subtitle: '직접방문', time: '11:00' },
    { id: '4', type: 'store', name: '최모리', subtitle: '강남', time: '17:00' },
  ],
  outstanding: [
    { id: 'a', name: '강미용실', phone: '010-1234-5678', amount: 500_000 },
    { id: 'b', name: '정헤어샵', phone: '010-9876-5432', amount: 300_000 },
  ],
  todos: [
    { id: 't1', text: '아임웹 상품 사진 교체' },
    { id: 't2', text: '신상품 일러스트 발주' },
    { id: 't3', text: '5월 결산 정리' },
  ],
  lowStock: 3,
  waybill: 35,
  newReviews: 7,
  purchasing: 2,
  supplies: 1,
  calendar: {
    5: [0, 1, 0], 8: [1, 0, 0], 12: [2, 0, 1], 15: [0, 0, 2],
    18: [1, 1, 0], 22: [0, 1, 0], 26: [3, 1, 1], 28: [2, 1, 0], 30: [0, 1, 0],
  } as Record<number, [number, number, number]>,
};

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

/* ──────────────────────────────────────────────────────────
 * 트렌디 컴포넌트 (마모루 모노크롬 베이스 + 절제된 상태색)
 * ──────────────────────────────────────────────────────── */

/** 1행 매출 KPI — 도넛 차트 + 3분할 미니 칩 */
function RevenueKPIDonut() {
  // SVG 도넛 차트 (외경 36 / 두께 6)
  const radius = 30;
  const circ = 2 * Math.PI * radius;
  const dash = (MOCK.pct / 100) * circ;

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-5 h-full flex flex-col">
      <div className="flex items-center justify-between mb-3">
        <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">이번달 매출</p>
        <button className="text-[10px] text-stone-400 hover:text-stone-700 transition">목표 수정</button>
      </div>

      <div className="flex items-center gap-4 flex-1">
        {/* 도넛 */}
        <div className="relative w-[88px] h-[88px] shrink-0">
          <svg viewBox="0 0 72 72" className="w-full h-full -rotate-90">
            <circle cx="36" cy="36" r={radius} fill="none" stroke="#F5F5F4" strokeWidth="6" />
            <circle
              cx="36" cy="36" r={radius} fill="none"
              stroke="#F59E0B" strokeWidth="6" strokeLinecap="round"
              strokeDasharray={`${dash} ${circ}`}
            />
          </svg>
          <div className="absolute inset-0 flex items-center justify-center">
            <span className="text-xl font-bold text-stone-900">{MOCK.pct}<span className="text-xs text-stone-500">%</span></span>
          </div>
        </div>

        {/* 수치 */}
        <div className="flex-1 min-w-0">
          <p className="text-lg font-bold text-stone-900 leading-tight">{fmtKRW(MOCK.current)}</p>
          <p className="text-[11px] text-stone-500 mt-0.5">목표 {fmtKRW(MOCK.monthGoal)}</p>
          <div className="grid grid-cols-3 gap-1 mt-3">
            <div className="text-center">
              <p className="text-[9px] text-stone-400">B2C</p>
              <p className="text-[11px] font-semibold text-stone-700">{fmtKRW(MOCK.b2c)}</p>
            </div>
            <div className="text-center border-l border-stone-100">
              <p className="text-[9px] text-stone-400">B2B</p>
              <p className="text-[11px] font-semibold text-stone-700">{fmtKRW(MOCK.b2b)}</p>
            </div>
            <div className="text-center border-l border-stone-100">
              <p className="text-[9px] text-stone-400">수리</p>
              <p className="text-[11px] font-semibold text-stone-700">{fmtKRW(MOCK.repair)}</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/** 4카드 공통 — 메인 수치 1개 크게 + 하단 보조 */
function CategoryCard({
  title, icon: Icon, mainValue, mainLabel, subStats, accent,
}: {
  title: string;
  icon: typeof ShoppingCart;
  mainValue: number;
  mainLabel: string;
  subStats?: string;
  accent: 'blue' | 'amber' | 'emerald' | 'stone';
}) {
  const accentColors = {
    blue:    'text-blue-600',
    amber:   'text-amber-600',
    emerald: 'text-emerald-600',
    stone:   'text-stone-800',
  };
  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full flex flex-col hover:border-stone-300 transition cursor-pointer group">
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center gap-1.5">
          <Icon size={13} className="text-stone-400" />
          <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">{title}</p>
        </div>
        <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
      </div>
      <div className="flex-1 flex flex-col justify-center">
        <p className={`text-3xl font-bold leading-none ${accentColors[accent]}`}>{mainValue}</p>
        <p className="text-[10px] text-stone-500 mt-1">{mainLabel}</p>
      </div>
      {subStats && (
        <p className="text-[10px] text-stone-400 mt-2 pt-2 border-t border-stone-100 truncate">{subStats}</p>
      )}
    </div>
  );
}

/** 1행 통합: 매출 KPI + 4카드 (한 줄) */
function Row1_RevenueAndCards() {
  return (
    <div className="grid grid-cols-12 gap-3">
      {/* 매출 KPI — 5/12 */}
      <div className="col-span-12 lg:col-span-5">
        <RevenueKPIDonut />
      </div>
      {/* 4카드 — 7/12, 내부 4분할 */}
      <div className="col-span-12 lg:col-span-7 grid grid-cols-2 sm:grid-cols-4 gap-3">
        <CategoryCard
          title="주문" icon={ShoppingCart} accent="blue"
          mainValue={MOCK.orders.payDone} mainLabel="결제완료"
          subStats={`준비 ${MOCK.orders.preparing} · 완료 ${MOCK.orders.delivered}`}
        />
        <CategoryCard
          title="상담" icon={MessageSquare} accent="amber"
          mainValue={MOCK.consultations.newIntake} mainLabel="신규접수"
          subStats={`예정 ${MOCK.consultations.confirmed} · 재요청 ${MOCK.consultations.needAction}`}
        />
        <CategoryCard
          title="복원수리" icon={Wrench} accent="emerald"
          mainValue={MOCK.repairs.readyToShip} mainLabel="출고대기"
          subStats={`이번달 ${MOCK.repairs.monthBags}자루`}
        />
        <CategoryCard
          title="제품 판매" icon={Store} accent="stone"
          mainValue={MOCK.sales.monthCount} mainLabel="이번달 판매"
          subStats={`B2C ${fmtKRW(MOCK.b2c)}`}
        />
      </div>
    </div>
  );
}

/** 2행 좌측: 트렌디 컴팩트 달력 */
function CompactCalendar() {
  const startOffset = 5;
  const totalDays = 31;
  const cells: (number | null)[] = [];
  for (let i = 0; i < startOffset; i++) cells.push(null);
  for (let d = 1; d <= totalDays; d++) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);

  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  const today = MOCK.todayDate.day;
  const selected = today;

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2">
          <span className="text-sm font-bold text-stone-900">2026년 5월</span>
        </div>
        <div className="flex items-center gap-1">
          <button className="w-7 h-7 rounded-lg hover:bg-stone-100 flex items-center justify-center transition" disabled>
            <ChevronLeft size={14} className="text-stone-500" />
          </button>
          <button className="w-7 h-7 rounded-lg hover:bg-stone-100 flex items-center justify-center transition" disabled>
            <ChevronRight size={14} className="text-stone-500" />
          </button>
        </div>
      </div>

      <div className="flex items-center gap-3 mb-2 text-[10px] text-stone-500">
        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />매장</span>
        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-violet-500" />출장</span>
        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-amber-500" />수리</span>
      </div>

      <div className="grid grid-cols-7 gap-1 mb-1">
        {dayNames.map((d, i) => (
          <div key={d} className={`text-center text-[10px] font-semibold py-1 ${i === 0 ? 'text-rose-500' : i === 6 ? 'text-blue-500' : 'text-stone-500'}`}>{d}</div>
        ))}
      </div>

      <div className="grid grid-cols-7 gap-1">
        {cells.map((d, idx) => {
          if (d === null) return <div key={idx} className="aspect-square" />;
          const events = MOCK.calendar[d];
          const isToday = d === today;
          const isSelected = d === selected;
          return (
            <div
              key={idx}
              className={`aspect-square rounded-lg p-1 text-[11px] cursor-pointer transition flex flex-col items-center justify-start ${
                isSelected ? 'bg-stone-900 text-white' :
                isToday ? 'bg-amber-50 text-amber-900 border border-amber-300' :
                'bg-white hover:bg-stone-50 text-stone-700'
              }`}
            >
              <div className="font-semibold leading-none mt-0.5">{d}</div>
              {events && (
                <div className="flex items-center gap-0.5 mt-1">
                  {events[0] > 0 && <span className={`w-1 h-1 rounded-full ${isSelected ? 'bg-emerald-300' : 'bg-emerald-500'}`} />}
                  {events[1] > 0 && <span className={`w-1 h-1 rounded-full ${isSelected ? 'bg-violet-300' : 'bg-violet-500'}`} />}
                  {events[2] > 0 && <span className={`w-1 h-1 rounded-full ${isSelected ? 'bg-amber-300' : 'bg-amber-500'}`} />}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/** 2행 우측: 선택일 타임라인 (기본 = 오늘) */
function DateTimeline() {
  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2">
          <span className="text-sm font-bold text-stone-900">
            {MOCK.todayDate.month}월 {MOCK.todayDate.day}일
            <span className="text-stone-400 text-xs font-normal ml-1.5">({MOCK.todayDate.weekday})</span>
          </span>
          <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-700 font-semibold">오늘</span>
        </div>
        <span className="text-[10px] text-stone-500">{MOCK.today.length}건</span>
      </div>

      {/* 타임라인 */}
      <div className="relative pl-4">
        {/* 세로 라인 */}
        <div className="absolute left-1 top-1.5 bottom-1.5 w-px bg-stone-200" />
        <div className="space-y-2.5">
          {MOCK.today.map((c) => {
            const config = c.type === 'store' ? { label: '매장', color: 'emerald', dotBg: 'bg-emerald-500', chipBg: 'bg-emerald-50 text-emerald-700' }
              : c.type === 'field' ? { label: '출장', color: 'violet', dotBg: 'bg-violet-500', chipBg: 'bg-violet-50 text-violet-700' }
              : { label: '수리', color: 'amber', dotBg: 'bg-amber-500', chipBg: 'bg-amber-50 text-amber-700' };
            return (
              <div key={c.id} className="relative flex items-center gap-2.5 group">
                {/* 점 */}
                <div className={`absolute -left-[15px] w-2.5 h-2.5 rounded-full ${config.dotBg} ring-2 ring-white`} />
                {/* 시간 */}
                <span className="text-xs font-semibold text-stone-500 w-10 shrink-0">{c.time}</span>
                {/* 칩 */}
                <span className={`text-[10px] px-1.5 py-0.5 rounded ${config.chipBg} font-semibold shrink-0`}>{config.label}</span>
                {/* 정보 */}
                <div className="flex-1 min-w-0 flex items-center justify-between gap-2">
                  <span className="text-xs font-medium text-stone-800 truncate">{c.name}</span>
                  <span className="text-[10px] text-stone-400 truncate">{c.subtitle}</span>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/** 3행 좌: 미수금 */
function OutstandingCard() {
  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-1.5">
          <AlertTriangle size={13} className="text-rose-500" />
          <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">미수금</p>
        </div>
        <span className="text-[10px] px-2 py-0.5 rounded-full bg-rose-50 text-rose-700 font-semibold">{MOCK.outstanding.length}건</span>
      </div>
      <div className="space-y-1.5">
        {MOCK.outstanding.map((c) => (
          <div key={c.id} className="flex items-center justify-between p-2 rounded-lg hover:bg-rose-50/40 cursor-pointer transition">
            <div className="min-w-0">
              <p className="text-xs font-semibold text-stone-800 truncate">{c.name}</p>
              <p className="text-[10px] text-stone-400 truncate">{c.phone}</p>
            </div>
            <span className="text-xs font-bold text-rose-600 shrink-0">{fmtKRW(c.amount)}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

/** 3행 중: 할일 메모 */
function TodoCard() {
  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full flex flex-col">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-1.5">
          <CheckCircle2 size={13} className="text-stone-400" />
          <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">할 일 메모</p>
        </div>
        <span className="text-[10px] px-2 py-0.5 rounded-full bg-stone-100 text-stone-600 font-semibold">{MOCK.todos.length}</span>
      </div>
      <div className="flex gap-1.5 mb-3">
        <input
          type="text"
          placeholder="할 일 입력..."
          className="flex-1 h-7 px-2.5 rounded-lg border border-stone-200 text-xs placeholder:text-stone-400 focus:outline-none focus:border-stone-400 transition"
          disabled
        />
        <button className="px-2 h-7 rounded-lg bg-stone-900 text-white text-[10px] font-semibold hover:bg-stone-800 transition flex items-center gap-0.5" disabled>
          <Plus size={11} />추가
        </button>
      </div>
      <div className="space-y-1 flex-1">
        {MOCK.todos.map((t) => (
          <div key={t.id} className="flex items-center gap-2 py-1 group">
            <span className="w-3.5 h-3.5 rounded border border-stone-300 shrink-0 group-hover:border-emerald-500 transition" />
            <span className="text-xs text-stone-700 flex-1">{t.text}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

/** 3행 우: 알림 5종 (가로 컴팩트) */
function AlertsCard() {
  const alerts = [
    { icon: PackageX,      label: '저재고',       count: MOCK.lowStock,    color: 'rose'    },
    { icon: Truck,         label: '운송장',       count: MOCK.waybill,     color: 'amber'   },
    { icon: Star,          label: '신규 후기',     count: MOCK.newReviews,  color: 'yellow'  },
    { icon: PackageOpen,   label: '매입 대기',     count: MOCK.purchasing,  color: 'blue'    },
    { icon: ClipboardList, label: '부자재',       count: MOCK.supplies,    color: 'stone'   },
  ] as const;

  const palette = {
    rose:   { text: 'text-rose-700',   bg: 'bg-rose-50',   icon: 'text-rose-500'   },
    amber:  { text: 'text-amber-700',  bg: 'bg-amber-50',  icon: 'text-amber-500'  },
    yellow: { text: 'text-yellow-700', bg: 'bg-yellow-50', icon: 'text-yellow-500' },
    blue:   { text: 'text-blue-700',   bg: 'bg-blue-50',   icon: 'text-blue-500'   },
    stone:  { text: 'text-stone-700',  bg: 'bg-stone-100', icon: 'text-stone-500'  },
  } as const;

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full">
      <div className="flex items-center justify-between mb-3">
        <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">시스템 알림</p>
      </div>
      <div className="grid grid-cols-5 gap-1.5">
        {alerts.map((a) => {
          const c = palette[a.color];
          return (
            <div key={a.label} className={`${c.bg} rounded-lg p-2 cursor-pointer hover:opacity-80 transition`}>
              <div className="flex items-center justify-between mb-1">
                <a.icon size={11} className={c.icon} />
                <span className={`text-xs font-bold ${c.text}`}>{a.count}</span>
              </div>
              <p className={`text-[9px] font-medium ${c.text} truncate`}>{a.label}</p>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 시안 B+ — 3행 압축 (최종 추천)
 * ──────────────────────────────────────────────────────── */
function SchemeBplus() {
  return (
    <div className="space-y-3">
      {/* 1행: 매출 KPI + 4카드 */}
      <Row1_RevenueAndCards />

      {/* 2행: 달력 + 선택일 타임라인 */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
        <CompactCalendar />
        <DateTimeline />
      </div>

      {/* 3행: 미수금 + 할일 + 알림 */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
        <OutstandingCard />
        <TodoCard />
        <AlertsCard />
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 페이지 본체
 * ──────────────────────────────────────────────────────── */
export default function DesignLabPage() {
  return (
    <>
      <Topbar title="🎨 디자인 모니터" />
      <div className="px-6 py-6 space-y-6 max-w-[1400px] mx-auto bg-stone-50 min-h-screen">
        {/* 헤더 */}
        <div className="px-5 py-4 rounded-2xl bg-stone-900 text-white">
          <div className="flex items-center gap-2 mb-1">
            <span className="text-base font-bold">🎨 TMS 디자인 모니터</span>
            <span className="text-[10px] px-2 py-0.5 rounded bg-white/15 uppercase tracking-wider">internal tool</span>
          </div>
          <p className="text-xs opacity-80 leading-relaxed">
            진행 중: <span className="font-semibold">§ 대시보드 리모델 — 시안 B+ (스크롤 최소화 + 트렌디)</span>
            <br />
            <span className="opacity-60">정적 mock · 운영 데이터 호출 X · 결정 후 § 즉시 삭제</span>
          </p>
        </div>

        {/* §: 시안 B+ */}
        <section className="space-y-5">
          <div className="flex items-center gap-3">
            <h2 className="text-lg font-bold text-stone-900">§ 대시보드 리모델 — 시안 B+ 최종안</h2>
            <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 font-semibold">진행 중</span>
          </div>

          {/* 4인 회의 결론 + 변경 요약 */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <div className="rounded-2xl border border-stone-200 bg-white p-4">
              <p className="text-[11px] font-bold text-stone-500 mb-2 uppercase tracking-wider">🧑‍💼 4인 회의 — B → B+ 진화</p>
              <ul className="text-xs text-stone-700 space-y-1.5 leading-relaxed">
                <li>🎨 <b>UX:</b> 1행 매출 + 4카드 통합 → 1초 테스트 통과율 ↑ (좌도넛 + 우4카드 골든레이쇼)</li>
                <li>💻 <b>개발:</b> grid-cols-12로 5:7 분할. SVG 도넛은 별도 패키지 불요</li>
                <li>🎯 <b>잡스:</b> 첫 화면에서 "돈·일정·할일" 3행으로 끝 — 스크롤 거의 없음</li>
                <li>🏢 <b>COO:</b> "오늘 일정" + 선택일 패널 통합 (기본값=오늘) → 위젯 1개 절감</li>
              </ul>
            </div>
            <div className="rounded-2xl border border-amber-300 bg-amber-50 p-4">
              <p className="text-[11px] font-bold text-amber-800 mb-2 uppercase tracking-wider">📐 B+ 핵심 변화</p>
              <ul className="text-xs text-amber-900 space-y-1.5 leading-relaxed">
                <li>✅ <b>5행 → 3행</b> 압축 (스크롤 ~60% 감소)</li>
                <li>✅ <b>매출 + 4카드 1행 통합</b> (좌 5/12 도넛 KPI / 우 7/12 4분할)</li>
                <li>✅ <b>"오늘 일정" 별도 카드 → 선택일 타임라인에 흡수</b> (기본값 오늘)</li>
                <li>✅ <b>알림 5종 → 우측 하단 가로 5분할 카드</b> (Footer 띠 X)</li>
                <li>✅ <b>SVG 도넛 + 컴팩트 카드 + 타임라인</b> — 트렌디 데이터 시각화</li>
              </ul>
            </div>
          </div>

          {/* 시안 B+ 본체 */}
          <div className="rounded-2xl border-2 border-stone-300 bg-stone-50 p-4">
            <SchemeBplus />
          </div>

          {/* 디자인 디테일 */}
          <div className="rounded-2xl border border-stone-200 bg-white p-5">
            <p className="text-[11px] font-bold text-stone-500 mb-3 uppercase tracking-wider">🎨 디자인 디테일 (TMS 가이드라인 — 마모루 컬러 베이스 + 트렌드)</p>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <p className="font-semibold text-stone-700 mb-1">컬러 시스템</p>
                <ul className="text-stone-500 space-y-0.5 text-[11px]">
                  <li>배경: stone-50</li>
                  <li>카드: white + stone-200 border</li>
                  <li>텍스트: stone-900 / 600 / 400</li>
                </ul>
              </div>
              <div>
                <p className="font-semibold text-stone-700 mb-1">상태색 (절제)</p>
                <ul className="text-stone-500 space-y-0.5 text-[11px]">
                  <li>매장 emerald · 출장 violet</li>
                  <li>수리 amber · 미수금 rose</li>
                  <li>신규 blue · 매출% amber</li>
                </ul>
              </div>
              <div>
                <p className="font-semibold text-stone-700 mb-1">레이아웃</p>
                <ul className="text-stone-500 space-y-0.5 text-[11px]">
                  <li>radius: 2xl (16px)</li>
                  <li>gap: 3 (12px)</li>
                  <li>패딩: 4~5 (16~20px)</li>
                </ul>
              </div>
              <div>
                <p className="font-semibold text-stone-700 mb-1">인터랙션</p>
                <ul className="text-stone-500 space-y-0.5 text-[11px]">
                  <li>hover: border 진하게</li>
                  <li>arrow icon translate-x</li>
                  <li>transition 표준 (200ms)</li>
                </ul>
              </div>
            </div>
          </div>

          {/* 회계 안전성 */}
          <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4">
            <p className="text-[11px] font-bold text-emerald-800 mb-2 uppercase tracking-wider">🔒 회계 안전성 검증</p>
            <ul className="text-xs text-emerald-800 space-y-1 leading-relaxed">
              <li>✅ 매출 합산 RPC (077·078·080·088) — 절대 미수정</li>
              <li>✅ useHubStats / useOutstandingAlert / useTodayConsultations — 호출 위치만 이동</li>
              <li>✅ KST 타임존 회귀 위험 0 — use-dashboard-stats.ts toLocalDateString 그대로</li>
              <li>✅ 4카드 데이터 정합성 — 단일 RPC 결과 독립 표시, 위치 이동만</li>
              <li>✅ 알림 5종 hook — 호출 위치만 이동</li>
              <li>⚠ "오늘 일정" → 선택일 패널 흡수 — useTodayConsultations 데이터 자체 무수정, 표시 컴포넌트만 통합</li>
            </ul>
          </div>

          {/* 결정 안내 */}
          <div className="rounded-2xl border-2 border-stone-900 bg-stone-900 text-white p-5 text-center">
            <p className="text-sm font-bold mb-1">📌 사장님 결정 대기 중</p>
            <p className="text-xs opacity-80 leading-relaxed">
              위 시안 B+를 실제 <code className="bg-white/10 px-1.5 py-0.5 rounded">/dashboard</code>에 적용하시겠습니까?
              <br />
              <span className="opacity-60">조정 사항(컬러·간격·위치 등)이 있으면 한 마디 해주세요. 적용 후 § 즉시 삭제됩니다.</span>
            </p>
          </div>
        </section>

        {/* 푸터 */}
        <div className="text-center text-[11px] text-stone-400 pt-4 border-t border-stone-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}
