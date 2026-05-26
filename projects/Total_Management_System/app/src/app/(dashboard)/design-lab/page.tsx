'use client';

/**
 * /design-lab — TMS 디자인 모니터 (사장님 + 클로드 협업 도구)
 *
 * 2026-05-26 사장님 운영 룰 (필독):
 *   ▶ 디자인 모니터는 그때그때 진행 중인 디자인 작업만 표시
 *   ▶ 작업 완료 (사장님 결정 + 실제 페이지 적용) 후 § 자동 삭제
 *   ▶ 영구 박제 X (실제 적용된 디자인은 운영 페이지 + memory/docs 에 박제)
 *   ▶ 새 작업 시작 시 § 신규 추가 → 비교 → 결정 → 적용 → § 삭제 (회전 도구)
 *
 * 현재 진행 작업: § 대시보드 리모델 — 상담 달력 통합 (3시안 비교)
 * 플랜: C:/Users/user/.claude/plans/tms-distributed-squid.md
 *
 * 운영 데이터 호출 X · 사이드바 메뉴 미노출 · URL 직접 접근만
 */

import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import {
  ShoppingCart, MessageSquare, Wrench, Store,
  Calendar, CheckCircle2, AlertTriangle, ClipboardList,
  PackageX, Truck, PackageOpen, Star,
  ChevronLeft, ChevronRight,
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
  orders: { payDone: 12, preparing: 3, shipping: 2, delivered: 8 },
  consultations: { newIntake: 5, confirmed: 8, needAction: 2 },
  repairs: { intakeNew: 3, pendingInbound: 2, workingCount: 4, readyToShip: 6 },
  sales: { monthCount: 32 },
  today: [
    { id: '1', type: 'store', name: '김미용', time: '14:00' },
    { id: '2', type: 'field', name: '박헤어', time: '15:30' },
    { id: '3', type: 'repair', name: '이살롱', time: '11:00' },
    { id: '4', type: 'store', name: '최모리', time: '17:00' },
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
  // 5월 달력 일정 (날짜 → [매장수, 출장수, 수리수])
  calendar: {
    5: [0, 1, 0], 8: [1, 0, 0], 12: [2, 0, 1], 15: [0, 0, 2],
    18: [1, 1, 0], 22: [0, 1, 0], 26: [3, 1, 1], 28: [2, 1, 0], 30: [0, 1, 0],
  } as Record<number, [number, number, number]>,
  selectedDate: 28,
};

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

/* ──────────────────────────────────────────────────────────
 * Mock 위젯 컴포넌트 (3시안 공통 재사용)
 * ──────────────────────────────────────────────────────── */

function MockRevenueKPI({ compact = false }: { compact?: boolean }) {
  return (
    <div className="bg-white rounded-lg border border-neutral-200 p-4">
      <div className="flex items-center justify-between mb-2">
        <p className="text-xs text-neutral-500">
          이번달 총매출 목표 <span className="text-neutral-400">(B2C + B2B + 복원수리)</span>
        </p>
      </div>
      <div className="flex items-end gap-3 mb-2">
        <span className={`${compact ? 'text-xl' : 'text-2xl'} font-bold text-yellow-600`}>{MOCK.pct}%</span>
        <span className="text-sm text-neutral-500">{fmtKRW(MOCK.current)} / {fmtKRW(MOCK.monthGoal)}</span>
      </div>
      <div className="w-full h-2.5 bg-neutral-100 rounded-full overflow-hidden">
        <div className="h-full rounded-full bg-yellow-500" style={{ width: `${MOCK.pct}%` }} />
      </div>
      <div className="mt-3 grid grid-cols-3 gap-2 text-center">
        <div className="rounded-lg bg-neutral-50 py-1.5">
          <p className="text-[10px] text-neutral-400">B2C 제품</p>
          <p className="text-xs font-bold text-neutral-700 mt-0.5">{fmtKRW(MOCK.b2c)}</p>
        </div>
        <div className="rounded-lg bg-neutral-50 py-1.5">
          <p className="text-[10px] text-neutral-400">B2B 제품</p>
          <p className="text-xs font-bold text-neutral-700 mt-0.5">{fmtKRW(MOCK.b2b)}</p>
        </div>
        <div className="rounded-lg bg-neutral-50 py-1.5">
          <p className="text-[10px] text-neutral-400">복원수리</p>
          <p className="text-xs font-bold text-neutral-700 mt-0.5">{fmtKRW(MOCK.repair)}</p>
        </div>
      </div>
    </div>
  );
}

function MockHubCard({
  title, icon: Icon, stats, summary, compact = false,
}: {
  title: string;
  icon: typeof ShoppingCart;
  stats: Array<{ label: string; value: number; color: string }>;
  summary?: string;
  compact?: boolean;
}) {
  return (
    <Card className={compact ? 'p-3' : ''}>
      <div className="flex items-center gap-2 mb-2">
        <Icon size={compact ? 14 : 16} className="text-neutral-600" />
        <span className={`${compact ? 'text-xs' : 'text-sm'} font-bold`}>{title}</span>
      </div>
      <div className={`grid gap-1 mb-1 ${stats.length === 1 ? 'grid-cols-1' : stats.length === 2 ? 'grid-cols-2' : stats.length === 3 ? 'grid-cols-3' : 'grid-cols-4'}`}>
        {stats.map((s) => (
          <div key={s.label} className="text-center">
            <p className="text-[10px] text-neutral-400">{s.label}</p>
            <p className={`${compact ? 'text-sm' : 'text-base'} font-bold ${s.color}`}>{s.value}</p>
          </div>
        ))}
      </div>
      {summary && <p className="text-[10px] text-neutral-500 truncate">{summary}</p>}
    </Card>
  );
}

function MockFourCards({ compact = false }: { compact?: boolean }) {
  return (
    <div className={`grid ${compact ? 'grid-cols-2 sm:grid-cols-4' : 'grid-cols-2'} gap-3`}>
      <MockHubCard title="주문" icon={ShoppingCart} compact={compact}
        stats={[
          { label: '결제완료', value: MOCK.orders.payDone, color: 'text-blue-600' },
          { label: '준비중', value: MOCK.orders.preparing, color: 'text-amber-600' },
        ]}
        summary={`이번달 ${fmtKRW(2_400_000)}`}
      />
      <MockHubCard title="상담" icon={MessageSquare} compact={compact}
        stats={[
          { label: '신규', value: MOCK.consultations.newIntake, color: 'text-amber-600' },
          { label: '예정', value: MOCK.consultations.confirmed, color: 'text-blue-600' },
        ]}
      />
      <MockHubCard title="복원수리" icon={Wrench} compact={compact}
        stats={[
          { label: '신규', value: MOCK.repairs.intakeNew, color: 'text-blue-600' },
          { label: '출고대기', value: MOCK.repairs.readyToShip, color: 'text-green-600' },
        ]}
        summary={`이번달 ${fmtKRW(MOCK.repair)} (15자루)`}
      />
      <MockHubCard title="제품 판매" icon={Store} compact={compact}
        stats={[
          { label: '이번달', value: MOCK.sales.monthCount, color: 'text-neutral-800' },
        ]}
        summary={`B2C ${fmtKRW(MOCK.b2c)} · B2B ${fmtKRW(MOCK.b2b)}`}
      />
    </div>
  );
}

function MockCalendar() {
  // 2026년 5월 (예시 — 31일까지, 5/1 = 금요일 → 시작 offset 5)
  const startOffset = 5;
  const totalDays = 31;
  const cells: (number | null)[] = [];
  for (let i = 0; i < startOffset; i++) cells.push(null);
  for (let d = 1; d <= totalDays; d++) cells.push(d);
  while (cells.length % 7 !== 0) cells.push(null);

  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  const today = 26;

  return (
    <Card>
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-2">
          <Calendar size={16} className="text-neutral-600" />
          <span className="text-sm font-bold">2026년 5월</span>
        </div>
        <div className="flex items-center gap-1">
          <button className="w-6 h-6 rounded hover:bg-neutral-100 flex items-center justify-center" disabled>
            <ChevronLeft size={14} />
          </button>
          <button className="w-6 h-6 rounded hover:bg-neutral-100 flex items-center justify-center" disabled>
            <ChevronRight size={14} />
          </button>
        </div>
      </div>
      {/* 범례 */}
      <div className="flex items-center gap-3 mb-2 text-[10px] text-neutral-500">
        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />매장</span>
        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-purple-500" />출장</span>
        <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-amber-500" />수리</span>
      </div>
      {/* 요일 */}
      <div className="grid grid-cols-7 gap-1 mb-1">
        {dayNames.map((d, i) => (
          <div key={d} className={`text-center text-[10px] font-semibold py-1 ${i === 0 ? 'text-red-500' : i === 6 ? 'text-blue-500' : 'text-neutral-500'}`}>{d}</div>
        ))}
      </div>
      {/* 셀 */}
      <div className="grid grid-cols-7 gap-1">
        {cells.map((d, idx) => {
          if (d === null) return <div key={idx} className="aspect-square" />;
          const events = MOCK.calendar[d];
          const isToday = d === today;
          const isSelected = d === MOCK.selectedDate;
          return (
            <div
              key={idx}
              className={`aspect-square rounded-md p-1 text-[11px] cursor-pointer transition border ${
                isSelected ? 'bg-neutral-900 text-white border-neutral-900' :
                isToday ? 'bg-amber-50 border-amber-300 text-amber-900' :
                'bg-white border-neutral-100 hover:bg-neutral-50'
              }`}
            >
              <div className="font-semibold leading-none">{d}</div>
              {events && (
                <div className="flex items-center gap-0.5 mt-1">
                  {events[0] > 0 && <span className="w-1 h-1 rounded-full bg-emerald-500" />}
                  {events[1] > 0 && <span className="w-1 h-1 rounded-full bg-purple-500" />}
                  {events[2] > 0 && <span className="w-1 h-1 rounded-full bg-amber-500" />}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </Card>
  );
}

function MockSelectedDatePanel() {
  return (
    <Card>
      <div className="flex items-center gap-2 mb-3">
        <Calendar size={16} className="text-neutral-600" />
        <span className="text-sm font-bold">5월 28일 일정</span>
        <Badge className="bg-neutral-100 text-neutral-600">3건</Badge>
      </div>
      <div className="space-y-2">
        <div className="flex items-center justify-between p-2 rounded-lg bg-emerald-50">
          <div className="flex items-center gap-2">
            <Badge className="bg-emerald-100 text-emerald-700 text-[10px]">매장</Badge>
            <span className="text-sm font-medium">김미용실</span>
          </div>
          <span className="text-xs text-neutral-500">14:00</span>
        </div>
        <div className="flex items-center justify-between p-2 rounded-lg bg-emerald-50">
          <div className="flex items-center gap-2">
            <Badge className="bg-emerald-100 text-emerald-700 text-[10px]">매장</Badge>
            <span className="text-sm font-medium">박헤어샵</span>
          </div>
          <span className="text-xs text-neutral-500">16:30</span>
        </div>
        <div className="flex items-center justify-between p-2 rounded-lg bg-purple-50">
          <div className="flex items-center gap-2">
            <Badge className="bg-purple-100 text-purple-700 text-[10px]">출장</Badge>
            <span className="text-sm font-medium">이살롱 (부산)</span>
          </div>
          <span className="text-xs text-neutral-500">10:00</span>
        </div>
      </div>
    </Card>
  );
}

function MockTodayConsults() {
  return (
    <Card>
      <div className="flex items-center gap-2 mb-3">
        <Calendar size={16} className="text-blue-500" />
        <span className="text-sm font-bold text-blue-700">오늘 일정 ({MOCK.today.length}건)</span>
      </div>
      <div className="space-y-1.5">
        {MOCK.today.map((c) => {
          const label = c.type === 'store' ? '매장' : c.type === 'field' ? '출장' : '수리';
          const color = c.type === 'store' ? 'bg-emerald-100 text-emerald-700'
            : c.type === 'field' ? 'bg-purple-100 text-purple-700'
            : 'bg-amber-100 text-amber-700';
          return (
            <div key={c.id} className="flex items-center justify-between p-2 rounded-lg hover:bg-neutral-50">
              <div className="flex items-center gap-2">
                <Badge className={`${color} text-[10px]`}>{label}</Badge>
                <span className="text-sm font-medium">{c.name}</span>
              </div>
              <span className="text-xs text-neutral-500">{c.time}</span>
            </div>
          );
        })}
      </div>
    </Card>
  );
}

function MockOutstanding() {
  return (
    <Card>
      <div className="flex items-center gap-2 mb-3">
        <AlertTriangle size={16} className="text-amber-500" />
        <span className="text-sm font-bold text-amber-700">미수금 ({MOCK.outstanding.length}건)</span>
      </div>
      <div className="space-y-1.5">
        {MOCK.outstanding.map((c) => (
          <div key={c.id} className="flex items-center justify-between p-2 rounded-lg hover:bg-amber-50">
            <div>
              <span className="text-sm font-medium">{c.name}</span>
              <span className="text-xs text-neutral-400 ml-2">{c.phone}</span>
            </div>
            <span className="text-sm font-bold text-amber-700">{fmtKRW(c.amount)}</span>
          </div>
        ))}
      </div>
    </Card>
  );
}

function MockTodo() {
  return (
    <Card>
      <div className="flex items-center gap-2 mb-3">
        <CheckCircle2 size={16} className="text-neutral-500" />
        <span className="text-sm font-bold">할 일 메모</span>
        <Badge className="bg-neutral-100 text-neutral-600">{MOCK.todos.length}</Badge>
      </div>
      <div className="flex gap-2 mb-3">
        <input
          type="text"
          placeholder="할 일 입력..."
          className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400"
          disabled
        />
        <button className="px-3 h-8 rounded-lg bg-neutral-900 text-white text-xs font-medium" disabled>추가</button>
      </div>
      <div className="space-y-1">
        {MOCK.todos.map((t) => (
          <div key={t.id} className="flex items-center gap-2 py-1">
            <span className="w-4 h-4 rounded border border-neutral-300 shrink-0" />
            <span className="text-sm text-neutral-700 flex-1">{t.text}</span>
          </div>
        ))}
      </div>
    </Card>
  );
}

function MockAlertsRow({ compact = true }: { compact?: boolean }) {
  const alerts = [
    { icon: PackageX, label: '저재고', count: MOCK.lowStock, color: 'text-red-600 bg-red-50' },
    { icon: Truck, label: '운송장 잔여', count: MOCK.waybill, color: 'text-amber-600 bg-amber-50' },
    { icon: Star, label: '신규 후기', count: MOCK.newReviews, color: 'text-yellow-600 bg-yellow-50' },
    { icon: PackageOpen, label: '매입 입고대기', count: MOCK.purchasing, color: 'text-indigo-600 bg-indigo-50' },
    { icon: ClipboardList, label: '부자재 주문필요', count: MOCK.supplies, color: 'text-neutral-600 bg-neutral-100' },
  ];
  if (compact) {
    return (
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-2">
        {alerts.map((a) => (
          <div key={a.label} className={`flex items-center gap-2 px-3 py-2 rounded-lg ${a.color}`}>
            <a.icon size={14} />
            <span className="text-xs font-medium flex-1 truncate">{a.label}</span>
            <span className="text-xs font-bold">{a.count}</span>
          </div>
        ))}
      </div>
    );
  }
  return (
    <div className="space-y-2">
      {alerts.map((a) => (
        <div key={a.label} className={`flex items-center gap-2 px-3 py-2 rounded-lg ${a.color}`}>
          <a.icon size={14} />
          <span className="text-xs font-medium flex-1">{a.label}</span>
          <span className="text-xs font-bold">{a.count}</span>
        </div>
      ))}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 시안 A — 사장님 시안 그대로 (3x2 균등 그리드)
 * ──────────────────────────────────────────────────────── */
function SchemeA() {
  return (
    <div className="space-y-4">
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <MockCalendar />
        <MockSelectedDatePanel />
        <div className="space-y-3">
          <MockRevenueKPI compact />
          <MockFourCards compact />
        </div>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <MockTodayConsults />
        <MockOutstanding />
        <MockTodo />
      </div>
      {/* 알림 5종 누락 경고 */}
      <div className="rounded-lg border-2 border-dashed border-red-300 bg-red-50 p-3 text-center">
        <p className="text-xs text-red-700 font-semibold">
          ⚠ 알림 5종 (저재고 / 운송장 / 신규 후기 / 매입 / 부자재) 누락 영역
          <br />
          <span className="font-normal text-red-600">현재 시안에는 표시 위치 없음 — 어디에 배치할지 결정 필요</span>
        </p>
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 시안 B — 추천: 매출 상단 풀폭 + 달력 중단 + 보조 + 알림 footer
 * ──────────────────────────────────────────────────────── */
function SchemeB() {
  return (
    <div className="space-y-4">
      {/* 1행: 매출 KPI 풀폭 */}
      <MockRevenueKPI />
      {/* 2행: 4카드 가로 풀폭 */}
      <MockFourCards compact />
      {/* 3행: 달력 + 선택일 패널 (2컬럼) */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <MockCalendar />
        <MockSelectedDatePanel />
      </div>
      {/* 4행: 보조 3카드 */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <MockTodayConsults />
        <MockOutstanding />
        <MockTodo />
      </div>
      {/* 5행: 알림 footer */}
      <div>
        <p className="text-[10px] uppercase tracking-wider text-neutral-400 mb-2">시스템 알림</p>
        <MockAlertsRow />
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 시안 C — 좌우 분할 (좌 2/3 달력 중심 / 우 1/3 매출+알림)
 * ──────────────────────────────────────────────────────── */
function SchemeC() {
  return (
    <div className="flex flex-col lg:flex-row gap-4">
      {/* 좌측 2/3: 달력 + 선택일 + 오늘일정 */}
      <div className="flex-1 min-w-0 space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <MockCalendar />
          <MockSelectedDatePanel />
        </div>
        <MockTodayConsults />
      </div>
      {/* 우측 1/3: 매출(컴팩트) + 4카드(2x2) + 미수금 + 할일 + 알림 */}
      <div className="w-full lg:w-[400px] shrink-0 space-y-3">
        <MockRevenueKPI compact />
        <MockFourCards compact />
        <MockOutstanding />
        <MockTodo />
        <MockAlertsRow compact={false} />
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 시안 비교 표
 * ──────────────────────────────────────────────────────── */
function ComparisonTable() {
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm border-collapse">
        <thead>
          <tr className="bg-neutral-100">
            <th className="border border-neutral-200 px-3 py-2 text-left">항목</th>
            <th className="border border-neutral-200 px-3 py-2">A안 (시안 그대로)</th>
            <th className="border border-neutral-200 px-3 py-2 bg-amber-50">B안 ⭐ 추천</th>
            <th className="border border-neutral-200 px-3 py-2">C안 (좌우 분할)</th>
          </tr>
        </thead>
        <tbody className="text-xs">
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">IA 시선 흐름</td>
            <td className="border border-neutral-200 px-3 py-2">균등 6분할 — 시선 분산</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50">매출→일정→액션→알림 (위계 명확)</td>
            <td className="border border-neutral-200 px-3 py-2">달력 중심 (사장님 시안 정신)</td>
          </tr>
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">1초 테스트</td>
            <td className="border border-neutral-200 px-3 py-2">⚠ 잔상 약함</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50">✅ 매출이 잔상으로 남음</td>
            <td className="border border-neutral-200 px-3 py-2">⚠ 달력만 잔상</td>
          </tr>
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">알림 5종 보존</td>
            <td className="border border-neutral-200 px-3 py-2">❌ 누락 (별도 결정 필요)</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50">✅ Footer 띠로 압축 보존</td>
            <td className="border border-neutral-200 px-3 py-2">✅ 우측 컬럼에 세로 보존</td>
          </tr>
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">매출 KPI 노출</td>
            <td className="border border-neutral-200 px-3 py-2">⚠ 우측 작게</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50">✅ 상단 풀폭 (사장님 첫 시선)</td>
            <td className="border border-neutral-200 px-3 py-2">⚠ 우측 컴팩트</td>
          </tr>
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">달력 크기</td>
            <td className="border border-neutral-200 px-3 py-2">중간 (1/3 폭)</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50">큼 (1/2 폭)</td>
            <td className="border border-neutral-200 px-3 py-2">중간 (좌측 1/2 폭)</td>
          </tr>
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">회계 영향도</td>
            <td className="border border-neutral-200 px-3 py-2 text-green-700">0 (UI만)</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50 text-green-700">0 (UI만)</td>
            <td className="border border-neutral-200 px-3 py-2 text-green-700">0 (UI만)</td>
          </tr>
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">작업량 (라인 추정)</td>
            <td className="border border-neutral-200 px-3 py-2">~150줄 + 알림 처리 미정</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50">~180줄 (제일 깔끔)</td>
            <td className="border border-neutral-200 px-3 py-2">~200줄 (좌우 균형 조절)</td>
          </tr>
          <tr>
            <td className="border border-neutral-200 px-3 py-2 font-semibold">모바일 반응형</td>
            <td className="border border-neutral-200 px-3 py-2">3컬럼 → 1컬럼 (긴 스크롤)</td>
            <td className="border border-neutral-200 px-3 py-2 bg-amber-50">자연스러운 수직 스택</td>
            <td className="border border-neutral-200 px-3 py-2">좌우 → 위아래 분할</td>
          </tr>
        </tbody>
      </table>
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
      <div className="px-6 py-6 space-y-8 max-w-[1400px] mx-auto">
        {/* 헤더 */}
        <div className="px-5 py-4 rounded-xl bg-neutral-900 text-white">
          <div className="flex items-center gap-2 mb-1">
            <span className="text-base font-bold">🎨 TMS 디자인 모니터</span>
            <span className="text-[10px] px-2 py-0.5 rounded bg-white/15 uppercase tracking-wider">internal tool</span>
          </div>
          <p className="text-xs opacity-80 leading-relaxed">
            진행 중: <span className="font-semibold">§ 대시보드 리모델 — 상담 달력 통합 (3시안 비교)</span>
            <br />
            <span className="opacity-60">정적 mock · 운영 데이터 호출 X · 결정 후 § 즉시 삭제</span>
          </p>
        </div>

        {/* §: 대시보드 리모델 비교 */}
        <section className="space-y-6">
          <div className="flex items-center gap-3">
            <h2 className="text-lg font-bold">§ 대시보드 리모델 — 상담 달력 통합</h2>
            <Badge className="bg-amber-100 text-amber-800 text-[10px]">진행 중</Badge>
          </div>

          {/* 4인 회의 요약 */}
          <div className="rounded-xl border border-neutral-200 bg-neutral-50 p-4">
            <p className="text-xs font-bold text-neutral-700 mb-2">🧑‍💼 4인 전문가 회의 결론</p>
            <ul className="text-xs text-neutral-600 space-y-1 leading-relaxed">
              <li>🎨 <b>UX:</b> 6분할 균등은 시각적 위계 약함. 매출 &gt; 일정 &gt; 액션 순으로 시선 흘러야</li>
              <li>💻 <b>개발:</b> ScheduleCalendar 컴포넌트 재사용 가능. useConsultations staleTime 60s 필수</li>
              <li>🎯 <b>잡스:</b> 사장님 매일 묻는 순서 — ①돈 들어왔나 ②오늘 누구 만나나 ③뭘 해야 하나</li>
              <li>🏢 <b>COO:</b> 알림 5종(저재고/운송장/리뷰/매입/부자재) 누락 위험 — Footer 보존 필요</li>
            </ul>
            <p className="text-xs text-amber-700 font-semibold mt-2">✅ 최종 추천: B안</p>
          </div>

          {/* 시안 A */}
          <div className="space-y-3">
            <div className="flex items-center gap-3 pt-2">
              <h3 className="text-base font-bold text-neutral-700">시안 A — 사장님 시안 그대로 (3x2 균등)</h3>
              <Badge className="bg-neutral-100 text-neutral-600 text-[10px]">원본</Badge>
            </div>
            <div className="rounded-xl border-2 border-neutral-200 bg-neutral-50 p-4">
              <SchemeA />
            </div>
          </div>

          {/* 시안 B (추천) */}
          <div className="space-y-3">
            <div className="flex items-center gap-3 pt-2">
              <h3 className="text-base font-bold text-amber-700">시안 B — 매출 상단 풀폭 + 달력 중단 + 알림 Footer</h3>
              <Badge className="bg-amber-100 text-amber-800 text-[10px]">⭐ 추천</Badge>
            </div>
            <div className="rounded-xl border-2 border-amber-300 bg-amber-50/30 p-4">
              <SchemeB />
            </div>
          </div>

          {/* 시안 C */}
          <div className="space-y-3">
            <div className="flex items-center gap-3 pt-2">
              <h3 className="text-base font-bold text-neutral-700">시안 C — 좌우 분할 (좌 2/3 달력 중심 / 우 1/3 매출)</h3>
              <Badge className="bg-neutral-100 text-neutral-600 text-[10px]">대안</Badge>
            </div>
            <div className="rounded-xl border-2 border-neutral-200 bg-neutral-50 p-4">
              <SchemeC />
            </div>
          </div>

          {/* 비교 표 */}
          <div className="space-y-3 pt-4">
            <h3 className="text-base font-bold text-neutral-700">📊 3시안 비교 표</h3>
            <ComparisonTable />
          </div>

          {/* 회계 안전성 */}
          <div className="rounded-xl border border-green-200 bg-green-50 p-4">
            <p className="text-xs font-bold text-green-700 mb-2">🔒 회계 안전성 검증 (3시안 공통)</p>
            <ul className="text-xs text-green-700 space-y-1 leading-relaxed">
              <li>✅ 매출 합산 RPC (077·078·080·088) — 절대 미수정</li>
              <li>✅ useHubStats / useOutstandingAlert / useTodayConsultations — 호출 위치만 이동</li>
              <li>✅ KST 타임존 회귀 위험 0 — use-dashboard-stats.ts 74행 toLocalDateString 그대로</li>
              <li>✅ 4카드 데이터 정합성 — 단일 RPC 결과 독립 표시, 위치 이동만</li>
              <li>✅ 알림 5종 hook — useLowStockAlert 등 5개 호출 위치만 이동</li>
              <li>⚠ 미수금 outstanding_balance 자동 갱신 로직 없음 — 기존 그대로 (이번 작업 범위 외)</li>
            </ul>
          </div>

          {/* 결정 안내 */}
          <div className="rounded-xl border-2 border-neutral-900 bg-neutral-900 text-white p-5 text-center">
            <p className="text-sm font-bold mb-1">📌 사장님 결정 대기 중</p>
            <p className="text-xs opacity-80">
              세 시안 중 채택안 선택 → 클로드가 실제 <code className="bg-white/10 px-1 rounded">/dashboard</code>에 적용 → § 즉시 삭제
              <br />
              <span className="opacity-60">"B안으로 가자" 또는 "A안 + 알림은 ○○에" 같은 조정안 모두 가능</span>
            </p>
          </div>
        </section>

        {/* 푸터 */}
        <div className="text-center text-[11px] text-neutral-400 pt-4 border-t border-neutral-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}
