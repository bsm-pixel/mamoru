'use client';

/**
 * /design-lab — TMS 디자인 모니터
 *
 * 진행 작업: § 복원수리 페이지군 톤 통일 — 매출 KPI 두 옵션(A:어두운 유지 / B:화이트 대시보드 일치) + 시안 B 톤 일관 적용
 * 운영 데이터 호출 X · 정적 mock
 */

import { Topbar } from '@/components/layout/topbar';
import {
  Inbox, Loader, CreditCard, AlertTriangle, TrendingUp,
  RefreshCw, ArrowRight, Wrench, Phone,
} from 'lucide-react';

const MOCK = {
  monthRepairAmount: 2_900_000,
  monthBags: 15,
  mamoru: { amount: 1_800_000, count: 9 },
  other:  { amount: 800_000,   count: 4 },
  b2b:    { amount: 300_000,   count: 2 },
  todayWork: { mamoru: 3, other: 1, count: 2 },
  weekWork: { mamoru: 8, other: 4, b2b: 2, count: 7 },
  stats: { intakeNew: 4, workingCount: 6, unpaidCount: 3, staleCount: 2 },
  repairs: [
    { id: '1', name: '김미용실', phone: '010-1234-5678', statusLabel: '신규접수', bar: 'bg-blue-500',    iconBg: 'bg-blue-50',    iconColor: 'text-blue-600',    chipBg: 'bg-blue-50 text-blue-700',       qty: 5, type: '마모루+타사' },
    { id: '2', name: '박헤어샵', phone: '010-9876-5432', statusLabel: '작업중',   bar: 'bg-amber-500',   iconBg: 'bg-amber-50',   iconColor: 'text-amber-600',   chipBg: 'bg-amber-50 text-amber-700',     qty: 3, type: '마모루' },
    { id: '3', name: '이살롱',   phone: '010-1111-2222', statusLabel: '출고대기', bar: 'bg-emerald-500', iconBg: 'bg-emerald-50', iconColor: 'text-emerald-600', chipBg: 'bg-emerald-50 text-emerald-700', qty: 8, type: '타사' },
  ],
};

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

/* ──────────────────────────────────────────────────────────
 * 매출 KPI 옵션 A: 어두운 (현재 + sales 페이지 패턴 유지)
 * ──────────────────────────────────────────────────────── */
function RevenueKPI_Dark() {
  return (
    <div className="rounded-2xl bg-stone-900 text-white overflow-hidden">
      <div className="flex items-center gap-3 px-5 py-4">
        <TrendingUp size={18} className="opacity-70" />
        <div>
          <p className="text-[11px] uppercase tracking-wider opacity-60 font-semibold">이번달 복원수리</p>
          <p className="text-2xl font-bold">{fmtKRW(MOCK.monthRepairAmount)}</p>
        </div>
        <p className="ml-auto text-xl font-bold">{MOCK.monthBags}<span className="text-xs opacity-60 ml-0.5">자루</span></p>
      </div>
      <div className="grid grid-cols-3 gap-px bg-white/10">
        {[
          { label: '마모루', ...MOCK.mamoru },
          { label: '타사',   ...MOCK.other  },
          { label: 'B2B',    ...MOCK.b2b    },
        ].map((s) => (
          <div key={s.label} className="px-3 py-2.5 text-center bg-stone-900">
            <p className="text-[10px] opacity-50 uppercase tracking-wider">{s.label}</p>
            <p className="text-sm font-bold mt-0.5">{fmtKRW(s.amount)}</p>
            <p className="text-[10px] opacity-40">{s.count}자루</p>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 매출 KPI 옵션 B: 화이트 (대시보드 톤 일치)
 * ──────────────────────────────────────────────────────── */
function RevenueKPI_White() {
  return (
    <div className="bg-white rounded-2xl border border-stone-200 overflow-hidden">
      <div className="flex items-center gap-3 px-5 py-4">
        <div className="w-10 h-10 rounded-xl bg-amber-50 flex items-center justify-center">
          <TrendingUp size={18} className="text-amber-600" />
        </div>
        <div>
          <p className="text-[11px] uppercase tracking-wider text-stone-500 font-semibold">이번달 복원수리</p>
          <p className="text-2xl font-bold text-stone-900">{fmtKRW(MOCK.monthRepairAmount)}</p>
        </div>
        <p className="ml-auto text-xl font-bold text-stone-900">{MOCK.monthBags}<span className="text-xs text-stone-500 ml-0.5">자루</span></p>
      </div>
      <div className="grid grid-cols-3 gap-px bg-stone-100">
        {[
          { label: '마모루', ...MOCK.mamoru },
          { label: '타사',   ...MOCK.other  },
          { label: 'B2B',    ...MOCK.b2b    },
        ].map((s) => (
          <div key={s.label} className="px-3 py-2.5 text-center bg-white">
            <p className="text-[10px] text-stone-400 uppercase tracking-wider">{s.label}</p>
            <p className="text-sm font-bold text-stone-700 mt-0.5">{fmtKRW(s.amount)}</p>
            <p className="text-[10px] text-stone-400">{s.count}자루</p>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 상태 4뱃지 (시안 B 톤 — CategoryCard)
 * ──────────────────────────────────────────────────────── */
function StatusCards() {
  const cards = [
    { label: '신규접수', value: MOCK.stats.intakeNew,    icon: Inbox,          accent: 'text-blue-600',   sub: '확인 필요' },
    { label: '진행중',   value: MOCK.stats.workingCount, icon: Loader,         accent: 'text-amber-600',  sub: '작업 중' },
    { label: '미입금',   value: MOCK.stats.unpaidCount,  icon: CreditCard,     accent: 'text-rose-600',   sub: '확인 필요' },
    { label: '3일경과',  value: MOCK.stats.staleCount,   icon: AlertTriangle,  accent: 'text-orange-600', sub: '지연' },
  ];
  return (
    <div className="grid grid-cols-4 gap-3">
      {cards.map((c) => {
        const Icon = c.icon;
        return (
          <button key={c.label} className="bg-white rounded-2xl border border-stone-200 p-4 hover:border-stone-300 transition group text-left">
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-1.5">
                <Icon size={13} className="text-stone-400" />
                <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">{c.label}</p>
              </div>
              <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
            </div>
            <p className={`text-3xl font-bold leading-none ${c.accent}`}>{c.value}</p>
            <p className="text-[10px] text-stone-500 mt-1">{c.sub}</p>
          </button>
        );
      })}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 6단계 탭 바 (시안 B 톤 — stone-900 활성)
 * ──────────────────────────────────────────────────────── */
function TabBar() {
  const tabs = [
    { key: 'intake',           label: '신규접수',     count: 4, active: false },
    { key: 'pickup_needed',    label: '수거접수필요', count: 1, active: false },
    { key: 'inbound_waiting',  label: '입고대기',     count: 2, active: false },
    { key: 'in_progress',      label: '진행중',       count: 6, active: true  },
    { key: 'ready_to_ship',    label: '출고대기',     count: 3, active: false },
    { key: 'shipped',          label: '출고완료',     count: 12,active: false },
  ];
  return (
    <div className="flex items-center gap-1 border-b border-stone-200 overflow-x-auto">
      {tabs.map((t) => (
        <button
          key={t.key}
          className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 transition whitespace-nowrap ${
            t.active ? 'border-stone-900 text-stone-900' : 'border-transparent text-stone-500 hover:text-stone-700'
          }`}
        >
          {t.label}
          <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-bold ${
            t.active ? 'bg-stone-900 text-white' : 'bg-stone-100 text-stone-500'
          }`}>{t.count}</span>
        </button>
      ))}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 리스트 행 (시안 B 톤 — 좌측 색 줄 + 아이콘 박스)
 * ──────────────────────────────────────────────────────── */
function RepairListPreview() {
  return (
    <div className="space-y-2">
      {MOCK.repairs.map((r) => (
        <div key={r.id} className="bg-white rounded-2xl border border-stone-200 overflow-hidden hover:border-stone-300 transition flex items-stretch group cursor-pointer">
          <div className={`w-1 ${r.bar}`} />
          <div className="flex-1 p-3 flex items-center gap-3 min-w-0">
            <div className={`w-10 h-10 rounded-lg ${r.iconBg} flex items-center justify-center shrink-0`}>
              <Wrench size={16} className={r.iconColor} />
            </div>
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 mb-0.5 flex-wrap">
                <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold shrink-0 ${r.chipBg}`}>{r.statusLabel}</span>
                <p className="text-sm font-semibold text-stone-800 truncate">{r.name}</p>
              </div>
              <div className="flex items-center gap-2 text-[11px] text-stone-500 flex-wrap">
                <span className="font-semibold text-stone-700">{r.qty}자루</span>
                <span className="text-stone-300">·</span>
                <span>{r.type}</span>
                <span className="text-stone-300">·</span>
                <span className="flex items-center gap-0.5"><Phone size={10} />{r.phone}</span>
              </div>
            </div>
            <ArrowRight size={14} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition shrink-0" />
          </div>
        </div>
      ))}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 시안 묶음 (옵션 A / B)
 * ──────────────────────────────────────────────────────── */
function SchemeWithOption(props: { kpi: 'dark' | 'white' }) {
  return (
    <div className="bg-stone-50 p-4 space-y-4">
      {/* 1행: 매출 KPI + 보조 카드 2개 */}
      <div className="grid grid-cols-1 lg:grid-cols-[1fr_auto_auto] gap-3">
        {props.kpi === 'dark' ? <RevenueKPI_Dark /> : <RevenueKPI_White />}
        <div className="bg-white rounded-2xl border border-stone-200 p-4 lg:w-44">
          <p className="text-[11px] uppercase tracking-wider text-stone-500 font-semibold mb-1">오늘 작업</p>
          <p className="text-2xl font-bold text-stone-900">{MOCK.todayWork.mamoru + MOCK.todayWork.other}<span className="text-xs text-stone-500 ml-0.5">정</span></p>
          <p className="text-[10px] text-stone-400 mt-0.5">마모루 {MOCK.todayWork.mamoru} · 타사 {MOCK.todayWork.other}</p>
        </div>
        <div className="bg-white rounded-2xl border border-stone-200 p-4 lg:w-44">
          <p className="text-[11px] uppercase tracking-wider text-stone-500 font-semibold mb-1">이번주 누적</p>
          <p className="text-2xl font-bold text-stone-900">{MOCK.weekWork.mamoru + MOCK.weekWork.other + MOCK.weekWork.b2b}<span className="text-xs text-stone-500 ml-0.5">정</span></p>
          <p className="text-[10px] text-stone-400 mt-0.5">마모루 {MOCK.weekWork.mamoru} · 타사 {MOCK.weekWork.other} · B2B {MOCK.weekWork.b2b}</p>
        </div>
      </div>

      {/* 2행: 상태 4카드 + 새로고침 */}
      <div className="grid grid-cols-12 gap-3">
        <div className="col-span-9">
          <StatusCards />
        </div>
        <div className="col-span-3">
          <button className="w-full h-full rounded-2xl border border-stone-200 bg-white hover:bg-stone-50 transition flex items-center justify-center gap-2 text-xs font-semibold text-stone-700 px-3 py-3">
            <RefreshCw size={14} />새로고침
          </button>
        </div>
      </div>

      {/* 3행: 6단계 탭 */}
      <TabBar />

      {/* 4행: 리스트 */}
      <RepairListPreview />
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
        <div className="px-5 py-4 rounded-2xl bg-stone-900 text-white">
          <div className="flex items-center gap-2 mb-1">
            <span className="text-base font-bold">🎨 TMS 디자인 모니터</span>
            <span className="text-[10px] px-2 py-0.5 rounded bg-white/15 uppercase tracking-wider">internal tool</span>
          </div>
          <p className="text-xs opacity-80 leading-relaxed">
            진행 중: <span className="font-semibold">§ 복원수리 페이지군 톤 통일 — 매출 KPI 2옵션(어두운/화이트) 비교</span>
            <br />
            <span className="opacity-60">정적 mock · 운영 데이터 호출 X · 결정 후 § 즉시 삭제</span>
          </p>
        </div>

        <section className="space-y-5">
          <div className="flex items-center gap-3">
            <h2 className="text-lg font-bold text-stone-900">§ 복원수리 페이지군 — 톤 통일 (매출 KPI 옵션 비교)</h2>
            <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 font-semibold">진행 중</span>
          </div>

          {/* 4인 회의 + 핵심 변경 */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <div className="rounded-2xl border border-stone-200 bg-white p-4">
              <p className="text-[11px] font-bold text-stone-500 mb-2 uppercase tracking-wider">🧑‍💼 4인 회의</p>
              <ul className="text-xs text-stone-700 space-y-1.5 leading-relaxed">
                <li>🎨 <b>UX:</b> 매출 KPI 색 결정이 핵심 — 어두운(임팩트) vs 화이트(대시보드 일치)</li>
                <li>💻 <b>개발:</b> 6단계 탭 + 상태 4뱃지 + 리스트 카드 일괄 모노크롬화. 데이터 hook 무수정</li>
                <li>🎯 <b>잡스:</b> 사장님 매일 6단계 확인 → 탭이 가장 자주 클릭. 활성 stone-900으로 절제</li>
                <li>🏢 <b>COO:</b> 미입금/3일경과 카드 항상 표시(0이면 회색) → 상태 일관성</li>
              </ul>
            </div>
            <div className="rounded-2xl border border-amber-300 bg-amber-50 p-4">
              <p className="text-[11px] font-bold text-amber-800 mb-2 uppercase tracking-wider">📐 공통 변경 (양 옵션 동일)</p>
              <ul className="text-xs text-amber-900 space-y-1.5 leading-relaxed">
                <li>✅ 페이지 배경 white → <b>stone-50</b></li>
                <li>✅ 상태 4뱃지 작은 버튼 → <b>CategoryCard 톤</b> (3xl 숫자 + uppercase)</li>
                <li>✅ 새로고침 → <b>우측 큰 카드 버튼</b></li>
                <li>✅ 6단계 탭 활성: <b>terracotta → stone-900</b></li>
                <li>✅ 리스트 행: <b>좌측 색 줄 + 아이콘 박스 + 화살표</b></li>
                <li>✅ 보조 카드(오늘작업/이번주누적): <b>rounded-2xl + uppercase 라벨</b></li>
              </ul>
            </div>
          </div>

          {/* 옵션 A — 어두운 매출 */}
          <div className="space-y-2">
            <div className="flex items-center gap-3">
              <h3 className="text-base font-bold text-stone-700">옵션 A — 매출 KPI 어두운 (현재 패턴 + sales 페이지 일치)</h3>
              <span className="text-[10px] px-2 py-0.5 rounded-full bg-stone-100 text-stone-600 font-semibold">유지</span>
            </div>
            <div className="rounded-2xl border-2 border-stone-200 overflow-hidden">
              <SchemeWithOption kpi="dark" />
            </div>
          </div>

          {/* 옵션 B — 화이트 매출 */}
          <div className="space-y-2">
            <div className="flex items-center gap-3">
              <h3 className="text-base font-bold text-stone-700">옵션 B — 매출 KPI 화이트 (대시보드 톤 일치)</h3>
              <span className="text-[10px] px-2 py-0.5 rounded-full bg-stone-100 text-stone-600 font-semibold">대시보드 일치</span>
            </div>
            <div className="rounded-2xl border-2 border-stone-200 overflow-hidden">
              <SchemeWithOption kpi="white" />
            </div>
          </div>

          {/* 비교 표 */}
          <div className="rounded-2xl border border-stone-200 bg-white p-4">
            <p className="text-[11px] font-bold text-stone-500 mb-3 uppercase tracking-wider">📊 옵션 A vs B 비교</p>
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr className="bg-stone-50">
                  <th className="border border-stone-200 px-3 py-2 text-left">항목</th>
                  <th className="border border-stone-200 px-3 py-2">A — 어두운</th>
                  <th className="border border-stone-200 px-3 py-2">B — 화이트</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="border border-stone-200 px-3 py-2 font-semibold">시각 임팩트</td>
                  <td className="border border-stone-200 px-3 py-2">✅ 강함 (눈에 띄게 매출 강조)</td>
                  <td className="border border-stone-200 px-3 py-2">⚪ 차분 (절제된 표현)</td>
                </tr>
                <tr>
                  <td className="border border-stone-200 px-3 py-2 font-semibold">대시보드 일치도</td>
                  <td className="border border-stone-200 px-3 py-2">⚪ 70% (보조 카드만 일치)</td>
                  <td className="border border-stone-200 px-3 py-2">✅ 100%</td>
                </tr>
                <tr>
                  <td className="border border-stone-200 px-3 py-2 font-semibold">sales 페이지 일치도</td>
                  <td className="border border-stone-200 px-3 py-2">✅ 100% (어두운 매출 카드)</td>
                  <td className="border border-stone-200 px-3 py-2">⚪ 50%</td>
                </tr>
                <tr>
                  <td className="border border-stone-200 px-3 py-2 font-semibold">변경 폭</td>
                  <td className="border border-stone-200 px-3 py-2">현재 유지 (neutral-900 → stone-900 미세 매칭)</td>
                  <td className="border border-stone-200 px-3 py-2">전체 화이트 리디자인</td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* 회계 안전성 */}
          <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4">
            <p className="text-[11px] font-bold text-emerald-800 mb-2 uppercase tracking-wider">🔒 회계·데이터 안전성</p>
            <ul className="text-xs text-emerald-800 space-y-1 leading-relaxed">
              <li>✅ 매출 합산 RPC (077·078·080·088) 무수정 — useRepairDashboardStats 결과 그대로 사용</li>
              <li>✅ 상태 전이 / 알림톡 / Google Calendar / 자동 배송 추적 무수정</li>
              <li>✅ 합포장 출고 / B2B 납품 연동 / 직접방문 시스템 무수정</li>
              <li>✅ 6단계 파이프라인 (status enum) 무수정 — 색만 통일</li>
              <li>⚠ REPAIR_STATUS_COLOR(format.ts) 검토 후 필요 시 1~2건 정규화</li>
            </ul>
          </div>

          {/* 결정 안내 */}
          <div className="rounded-2xl border-2 border-stone-900 bg-stone-900 text-white p-5 text-center">
            <p className="text-sm font-bold mb-1">📌 사장님 결정 대기 중</p>
            <p className="text-xs opacity-80 leading-relaxed">
              <b>옵션 A (어두운 유지)</b> / <b>옵션 B (화이트 일치)</b> 중 선택 → 클로드가 실제 <code className="bg-white/10 px-1.5 py-0.5 rounded">/repairs</code> + 컴포넌트에 적용
              <br />
              <span className="opacity-60">"A로 가자" 또는 "B로 가자" 한 마디. 적용 후 § 즉시 삭제 + push.</span>
            </p>
          </div>
        </section>

        <div className="text-center text-[11px] text-stone-400 pt-4 border-t border-stone-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}
