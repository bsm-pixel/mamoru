'use client';

/**
 * /design-lab — TMS 디자인 모니터
 *
 * 진행 작업: § 어두운 매출 카드 톤 비교 (4안: 원본 stone-900 / 그라데이션 / stone-800 미디엄 / 노이즈+그라데이션)
 *   → 복원수리 + sales 페이지 동시 적용 예정
 * 운영 데이터 호출 X · 정적 mock
 */

import { Topbar } from '@/components/layout/topbar';
import { TrendingUp } from 'lucide-react';

const MOCK = {
  total: 11_200_000,
  bags: 32,
  mamoru: { amount: 5_400_000, count: 18 },
  other:  { amount: 3_600_000, count: 10 },
  b2b:    { amount: 2_200_000, count: 4 },
};

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

/* ──────────────────────────────────────────────────────────
 * 매출 카드 4가지 톤 (옵션 A의 부드러움 변형)
 * ──────────────────────────────────────────────────────── */

/** A1 — 현재(stone-900 단색): 강함, "전광판" */
function CardA1() {
  return (
    <div className="rounded-2xl bg-stone-900 text-white overflow-hidden">
      <CardHeader />
      <CardSplit divider="bg-white/10" cellBg="bg-stone-900" />
    </div>
  );
}

/** A2 — 그라데이션 (stone-800 → stone-900): 자연스러운 깊이감, 트렌디 */
function CardA2() {
  return (
    <div className="rounded-2xl bg-gradient-to-br from-stone-800 to-stone-900 text-white overflow-hidden ring-1 ring-white/5">
      <CardHeader />
      <CardSplit divider="bg-white/10" cellBg="bg-stone-900/40" />
    </div>
  );
}

/** A3 — stone-800 단색 + 미세 보더: 한 톤 부드러움 */
function CardA3() {
  return (
    <div className="rounded-2xl bg-stone-800 text-white overflow-hidden ring-1 ring-white/5">
      <CardHeader />
      <CardSplit divider="bg-white/8" cellBg="bg-stone-800" />
    </div>
  );
}

/** A4 — 차콜 그라데이션 + 살짝 푸르스름 (트렌드 차콜): 가장 모던, 부드러움 */
function CardA4() {
  return (
    <div className="rounded-2xl bg-gradient-to-br from-zinc-800 via-stone-800 to-stone-900 text-white overflow-hidden ring-1 ring-white/5 shadow-lg shadow-stone-900/10">
      <CardHeader textOpacity="opacity-95" />
      <CardSplit divider="bg-white/10" cellBg="bg-zinc-900/30" />
    </div>
  );
}

function CardHeader({ textOpacity = '' }: { textOpacity?: string }) {
  return (
    <div className={`flex items-center gap-3 px-5 py-4 ${textOpacity}`}>
      <TrendingUp size={18} className="opacity-60" />
      <div>
        <p className="text-[11px] uppercase tracking-wider opacity-60 font-semibold">이번달 복원수리</p>
        <p className="text-2xl font-bold">{fmtKRW(MOCK.total)}</p>
      </div>
      <p className="ml-auto text-xl font-bold">
        {MOCK.bags}<span className="text-xs opacity-60 ml-0.5">자루</span>
      </p>
    </div>
  );
}

function CardSplit({ divider, cellBg }: { divider: string; cellBg: string }) {
  return (
    <div className={`grid grid-cols-3 gap-px ${divider}`}>
      {[
        { label: '마모루', ...MOCK.mamoru },
        { label: '타사',   ...MOCK.other  },
        { label: 'B2B',    ...MOCK.b2b    },
      ].map((s) => (
        <div key={s.label} className={`px-3 py-2.5 text-center ${cellBg}`}>
          <p className="text-[10px] opacity-50 uppercase tracking-wider">{s.label}</p>
          <p className="text-sm font-bold mt-0.5">{fmtKRW(s.amount)}</p>
          <p className="text-[10px] opacity-40">{s.count}자루</p>
        </div>
      ))}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 페이지 본체
 * ──────────────────────────────────────────────────────── */
export default function DesignLabPage() {
  const options = [
    { key: 'A1', label: 'A1 — 원본 (stone-900 단색)',                 sub: '강함, "전광판" 느낌',                   Card: CardA1, badge: '현재' },
    { key: 'A2', label: 'A2 — 그라데이션 (stone-800 → stone-900)',    sub: '자연 깊이감, 트렌디, 가장 균형',         Card: CardA2, badge: '⭐ 추천' },
    { key: 'A3', label: 'A3 — stone-800 단색',                        sub: '한 톤 부드러움, 평평한 톤',              Card: CardA3, badge: '대안 1' },
    { key: 'A4', label: 'A4 — 차콜 그라데이션 (zinc + stone)',         sub: '가장 모던, 미세 푸르스름 트렌드 차콜',   Card: CardA4, badge: '대안 2' },
  ];

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
            진행 중: <span className="font-semibold">§ 어두운 매출 카드 톤 비교 (옵션 A 부드러운 변형)</span>
            <br />
            <span className="opacity-60">사장님 피드백: "땅땅 박혀있는 전광판 같아서 살짝 은은하게" — sales/repairs 동시 적용 예정</span>
          </p>
        </div>

        <section className="space-y-5">
          <div className="flex items-center gap-3">
            <h2 className="text-lg font-bold text-stone-900">§ 어두운 매출 카드 톤 비교 (A1 → A4)</h2>
            <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 font-semibold">진행 중</span>
          </div>

          {/* 4인 회의 */}
          <div className="rounded-2xl border border-stone-200 bg-white p-4">
            <p className="text-[11px] font-bold text-stone-500 mb-2 uppercase tracking-wider">🧑‍💼 4인 회의</p>
            <ul className="text-xs text-stone-700 space-y-1.5 leading-relaxed">
              <li>🎨 <b>UX:</b> 단색 검정(A1)은 stone-50 배경과 contrast 너무 강함 → 그라데이션(A2)이 시선 자연 흐름</li>
              <li>💻 <b>개발:</b> 그라데이션은 Tailwind 단일 클래스. 추가 비용 0. A2 = gradient-to-br + ring-1 white/5</li>
              <li>🎯 <b>잡스:</b> "은은하게"의 핵심 = 카드와 배경 사이 단절감 제거. ring으로 미세 윤곽 + 그라데이션으로 깊이</li>
              <li>🏢 <b>COO:</b> 모든 어두운 매출 카드 = 통일 → sales/repairs/향후 매입 모두 같은 톤</li>
            </ul>
            <p className="text-xs text-amber-700 font-semibold mt-2">✅ 추천: A2 (그라데이션) — 부드러움 + 임팩트 + 트렌디 균형</p>
          </div>

          {/* 4가지 옵션 동시 비교 */}
          <div className="space-y-4">
            {options.map((opt) => {
              const Card = opt.Card;
              return (
                <div key={opt.key} className="space-y-2">
                  <div className="flex items-center gap-3">
                    <h3 className="text-sm font-bold text-stone-700">{opt.label}</h3>
                    <span className={`text-[10px] px-2 py-0.5 rounded-full font-semibold ${
                      opt.badge === '⭐ 추천' ? 'bg-amber-100 text-amber-800' : 'bg-stone-100 text-stone-600'
                    }`}>{opt.badge}</span>
                  </div>
                  <p className="text-[11px] text-stone-500">{opt.sub}</p>
                  <div className="max-w-[640px]">
                    <Card />
                  </div>
                </div>
              );
            })}
          </div>

          {/* 적용 범위 안내 */}
          <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4">
            <p className="text-[11px] font-bold text-emerald-800 mb-2 uppercase tracking-wider">📍 채택안 적용 범위</p>
            <ul className="text-xs text-emerald-800 space-y-1 leading-relaxed">
              <li>✅ <b>/repairs</b> 상단 매출 KPI</li>
              <li>✅ <b>/sales</b> 매출 카드 (안 3 — 기존 어두운 카드 동일 톤 변경)</li>
              <li>✅ 향후 <b>/purchasing</b> 매입 합계 카드 (예정)</li>
              <li>📌 "매출/돈 합계 = 어두운 카드 (부드러운 톤)" 규칙 정립</li>
            </ul>
          </div>

          {/* 결정 안내 */}
          <div className="rounded-2xl border-2 border-stone-900 bg-stone-900 text-white p-5 text-center">
            <p className="text-sm font-bold mb-1">📌 사장님 결정 대기 중</p>
            <p className="text-xs opacity-80 leading-relaxed">
              <b>A1 / A2 / A3 / A4</b> 중 하나 선택 → 클로드가 /repairs + /sales 동시 적용
              <br />
              <span className="opacity-60">"A2로 가자" 한 마디. 적용 후 § 즉시 삭제 + push.</span>
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
