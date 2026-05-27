'use client';

/**
 * /design-lab — TMS 디자인 모니터
 *
 * 진행 작업: § 상담 페이지군 톤 통일 (시안 A/B 비교 → 추천 B)
 * 운영 데이터 호출 X · 정적 mock
 */

import { Topbar } from '@/components/layout/topbar';
import {
  Store, Truck, MessageCircle, Inbox, Loader, CheckCircle,
  RefreshCw, CalendarPlus, ArrowRight, Phone, MapPin,
} from 'lucide-react';

/* ──────────────────────────────────────────────────────────
 * Mock 데이터
 * ──────────────────────────────────────────────────────── */
const MOCK = {
  stats: { newIntake: 7, inProgress: 12, completedMonth: 38 },
  needAction: { store_visit: 0, field_request: 3, talk_consult: 2 },
  consults: [
    { id: '1', name: '김미용실', phone: '010-1234-5678', date: '5월 28일 (목)', time: '14:00', region: '강남' },
    { id: '2', name: '박헤어샵', phone: '010-9876-5432', date: '5월 28일 (목)', time: '16:30', region: '강남' },
    { id: '3', name: '이살롱', phone: '010-1111-2222', date: '5월 30일 (토)', time: '11:00', region: '부산' },
  ],
};

/* ──────────────────────────────────────────────────────────
 * 시안 A — 보수적: 요약 카드 + 탭만 모노크롬화 (배경 white 유지)
 * ──────────────────────────────────────────────────────── */
function SchemeA() {
  return (
    <div className="bg-white px-4 py-4 space-y-4">
      {/* 상단: 요약 카드 + 새로고침 */}
      <div className="flex items-center gap-3">
        <div className="flex gap-2 flex-1 min-w-0">
          {/* 보수: bg-white + border, 텍스트만 강조색 */}
          <button className="flex items-center gap-2 px-3 py-2 rounded-lg border border-stone-200 hover:bg-stone-50 transition">
            <Inbox size={14} className="text-blue-600 shrink-0" />
            <span className="text-xs text-stone-500">신규</span>
            <span className="text-sm font-bold text-blue-700">{MOCK.stats.newIntake}</span>
          </button>
          <button className="flex items-center gap-2 px-3 py-2 rounded-lg border border-stone-200 hover:bg-stone-50 transition">
            <Loader size={14} className="text-amber-600 shrink-0" />
            <span className="text-xs text-stone-500">진행</span>
            <span className="text-sm font-bold text-amber-700">{MOCK.stats.inProgress}</span>
          </button>
          <button className="flex items-center gap-2 px-3 py-2 rounded-lg border border-stone-200 hover:bg-stone-50 transition">
            <CheckCircle size={14} className="text-emerald-600 shrink-0" />
            <span className="text-xs text-stone-500">완료</span>
            <span className="text-sm font-bold text-emerald-700">{MOCK.stats.completedMonth}</span>
          </button>
        </div>
        <button className="px-3 py-1.5 rounded-lg border border-stone-200 text-xs font-semibold text-stone-700 hover:bg-stone-50 transition flex items-center gap-1.5">
          <RefreshCw size={14} />새로고침
        </button>
      </div>

      {/* 3탭 — 활성: stone-900 베이스 */}
      <div className="flex items-center justify-between border-b border-stone-200">
        <div className="flex gap-1">
          {[
            { key: 'store_visit', label: '매장방문', icon: Store, active: true },
            { key: 'field_request', label: '출장요청', icon: Truck, badge: 3 },
            { key: 'talk_consult', label: '온라인상담', icon: MessageCircle, badge: 2 },
          ].map((tab) => {
            const Icon = tab.icon;
            return (
              <button
                key={tab.key}
                className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 transition relative ${
                  tab.active ? 'border-stone-900 text-stone-900' : 'border-transparent text-stone-500 hover:text-stone-700'
                }`}
              >
                <Icon size={14} />
                {tab.label}
                {tab.badge && (
                  <span className="ml-1 inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-rose-500 text-white text-[10px] font-bold leading-none">
                    {tab.badge}
                  </span>
                )}
              </button>
            );
          })}
        </div>
        <button className="px-3 py-1.5 rounded-lg bg-stone-900 text-white text-xs font-semibold mb-1 flex items-center gap-1.5 hover:bg-stone-800 transition">
          <CalendarPlus size={14} />일정수동등록
        </button>
      </div>

      {/* 리스트 (간단 mock) */}
      <div className="space-y-2">
        {MOCK.consults.map((c) => (
          <div key={c.id} className="bg-white rounded-2xl border border-stone-200 p-3 flex items-center gap-3 hover:border-stone-300 transition">
            <div className="w-10 h-10 rounded-lg bg-emerald-50 flex items-center justify-center shrink-0">
              <Store size={16} className="text-emerald-600" />
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-sm font-semibold text-stone-800 truncate">{c.name}</p>
              <p className="text-xs text-stone-500 truncate">{c.date} · {c.time} · {c.region}</p>
            </div>
            <span className="text-[10px] px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-semibold shrink-0">확정</span>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 시안 B (추천) — 적극: 배경 stone-50 + 대시보드 CategoryCard 톤 요약 + 모노크롬 탭
 * ──────────────────────────────────────────────────────── */
function SchemeB() {
  return (
    <div className="bg-stone-50 px-4 py-4 space-y-4">
      {/* 1행: 요약 3카드 (대시보드 CategoryCard 톤) + 새로고침 */}
      <div className="grid grid-cols-12 gap-3">
        <div className="col-span-9 grid grid-cols-3 gap-3">
          {[
            { label: '신규', value: MOCK.stats.newIntake, sub: '확인 필요', icon: Inbox, accent: 'text-blue-600' },
            { label: '진행', value: MOCK.stats.inProgress, sub: '일정 조율 중', icon: Loader, accent: 'text-amber-600' },
            { label: '완료', value: MOCK.stats.completedMonth, sub: '이번달', icon: CheckCircle, accent: 'text-emerald-600' },
          ].map((s) => {
            const Icon = s.icon;
            return (
              <button
                key={s.label}
                className="bg-white rounded-2xl border border-stone-200 p-4 hover:border-stone-300 transition group text-left"
              >
                <div className="flex items-center justify-between mb-2">
                  <div className="flex items-center gap-1.5">
                    <Icon size={13} className="text-stone-400" />
                    <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">{s.label}</p>
                  </div>
                  <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
                </div>
                <p className={`text-3xl font-bold leading-none ${s.accent}`}>{s.value}</p>
                <p className="text-[10px] text-stone-500 mt-1">{s.sub}</p>
              </button>
            );
          })}
        </div>
        <div className="col-span-3 flex items-center">
          <button className="w-full h-full rounded-2xl border border-stone-200 bg-white hover:bg-stone-50 transition flex items-center justify-center gap-2 text-xs font-semibold text-stone-700">
            <RefreshCw size={14} />새로고침
          </button>
        </div>
      </div>

      {/* 2행: 3탭 + 일정수동등록 */}
      <div className="flex items-center justify-between border-b border-stone-200">
        <div className="flex gap-1">
          {[
            { key: 'store_visit', label: '매장방문', icon: Store, active: true },
            { key: 'field_request', label: '출장요청', icon: Truck, badge: 3 },
            { key: 'talk_consult', label: '온라인상담', icon: MessageCircle, badge: 2 },
          ].map((tab) => {
            const Icon = tab.icon;
            return (
              <button
                key={tab.key}
                className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 transition relative ${
                  tab.active ? 'border-stone-900 text-stone-900' : 'border-transparent text-stone-500 hover:text-stone-700'
                }`}
              >
                <Icon size={14} />
                {tab.label}
                {tab.badge && (
                  <span className="ml-1 inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-rose-500 text-white text-[10px] font-bold leading-none">
                    {tab.badge}
                  </span>
                )}
              </button>
            );
          })}
        </div>
        <button className="px-3 py-1.5 rounded-lg bg-stone-900 text-white text-xs font-semibold mb-1 flex items-center gap-1.5 hover:bg-stone-800 transition">
          <CalendarPlus size={14} />일정수동등록
        </button>
      </div>

      {/* 3행: 리스트 — 좌측 색 줄 + 트렌디 카드 */}
      <div className="space-y-2">
        {MOCK.consults.map((c) => (
          <div key={c.id} className="bg-white rounded-2xl border border-stone-200 overflow-hidden hover:border-stone-300 transition flex items-stretch group cursor-pointer">
            {/* 좌측 색 줄 (상태별) */}
            <div className="w-1 bg-emerald-500" />
            <div className="flex-1 p-3 flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-emerald-50 flex items-center justify-center shrink-0">
                <Store size={16} className="text-emerald-600" />
              </div>
              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2 mb-0.5">
                  <p className="text-sm font-semibold text-stone-800 truncate">{c.name}</p>
                  <span className="text-[10px] px-1.5 py-0.5 rounded bg-emerald-50 text-emerald-700 font-semibold">확정</span>
                </div>
                <div className="flex items-center gap-2 text-[11px] text-stone-500">
                  <span>{c.date}</span>
                  <span className="text-stone-300">·</span>
                  <span className="font-semibold text-stone-700">{c.time}</span>
                  <span className="text-stone-300">·</span>
                  <span className="flex items-center gap-0.5"><MapPin size={10} />{c.region}</span>
                </div>
              </div>
              <span className="text-xs text-stone-500 truncate hidden sm:inline">{c.phone}</span>
              <ArrowRight size={14} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition shrink-0" />
            </div>
          </div>
        ))}
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
            진행 중: <span className="font-semibold">§ 상담 페이지군 톤 통일 — 시안 A(보수) vs B(적극·추천)</span>
            <br />
            <span className="opacity-60">정적 mock · 운영 데이터 호출 X · 결정 후 § 즉시 삭제</span>
          </p>
        </div>

        {/* §: 상담 페이지군 톤 통일 */}
        <section className="space-y-5">
          <div className="flex items-center gap-3">
            <h2 className="text-lg font-bold text-stone-900">§ 상담 페이지군 — 톤 통일 시안 비교</h2>
            <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 font-semibold">진행 중</span>
          </div>

          {/* 4인 회의 + 변경 영역 요약 */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <div className="rounded-2xl border border-stone-200 bg-white p-4">
              <p className="text-[11px] font-bold text-stone-500 mb-2 uppercase tracking-wider">🧑‍💼 4인 회의</p>
              <ul className="text-xs text-stone-700 space-y-1.5 leading-relaxed">
                <li>🎨 <b>UX:</b> 요약 3카드 + 탭이 시안 B+와 가장 큰 차이 — CategoryCard 톤 그대로 가져오자</li>
                <li>💻 <b>개발:</b> 주 수정 = page.tsx + store-visit-list.tsx + field-request-list.tsx (lib/utils/format.ts orange-100 1줄 normalize)</li>
                <li>🎯 <b>잡스:</b> 사장님 매일 들어가는 페이지 → 1초 테스트로 신규/진행/완료 즉시 파악</li>
                <li>🏢 <b>COO:</b> 탭 전환 부드럽게, 색 절제 — terracotta → stone-900</li>
              </ul>
            </div>
            <div className="rounded-2xl border border-amber-300 bg-amber-50 p-4">
              <p className="text-[11px] font-bold text-amber-800 mb-2 uppercase tracking-wider">📐 변경 영역 (B안 기준)</p>
              <ul className="text-xs text-amber-900 space-y-1.5 leading-relaxed">
                <li>✅ 페이지 배경 white → <b>stone-50</b> (대시보드 톤 일치)</li>
                <li>✅ 요약 3카드 작은 버튼 → <b>대시보드 CategoryCard 톤</b> (큰 숫자 + uppercase 라벨)</li>
                <li>✅ 탭 활성: <b>terracotta → stone-900</b></li>
                <li>✅ 리스트 행: <b>좌측 색 줄 + 좌측 아이콘 박스 + 우측 화살표</b> (sales 페이지 패턴)</li>
                <li>✅ 일정수동등록 버튼: <b>secondary → stone-900 검정</b></li>
              </ul>
            </div>
          </div>

          {/* 시안 A */}
          <div className="space-y-2">
            <div className="flex items-center gap-3">
              <h3 className="text-base font-bold text-stone-700">시안 A — 보수적 (요약 카드 + 탭만 모노크롬)</h3>
              <span className="text-[10px] px-2 py-0.5 rounded-full bg-stone-100 text-stone-600 font-semibold">최소 변경</span>
            </div>
            <div className="rounded-2xl border-2 border-stone-200 bg-white overflow-hidden">
              <SchemeA />
            </div>
            <ul className="text-[11px] text-stone-500 list-disc list-inside ml-1 leading-relaxed">
              <li>장점: 변경 범위 작음 (~20줄), 회귀 위험 최소</li>
              <li>단점: 대시보드와 톤 일치도 70% — 페이지 진입 시 살짝 다른 느낌</li>
            </ul>
          </div>

          {/* 시안 B (추천) */}
          <div className="space-y-2">
            <div className="flex items-center gap-3">
              <h3 className="text-base font-bold text-amber-700">시안 B — 적극적 (대시보드 톤 완전 일치)</h3>
              <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 font-semibold">⭐ 추천</span>
            </div>
            <div className="rounded-2xl border-2 border-amber-300 bg-amber-50/30 overflow-hidden">
              <SchemeB />
            </div>
            <ul className="text-[11px] text-stone-500 list-disc list-inside ml-1 leading-relaxed">
              <li>장점: 대시보드 완전 톤 일치, 1초 테스트 통과, 정보 위계 명확</li>
              <li>단점: 변경 범위 큼 (~80줄, 3파일) — 회귀 점검 더 꼼꼼히 필요</li>
              <li>리스트 행은 sales 페이지 패턴(좌측 색 줄 + 도트) 차용 → TMS 전체 통일성</li>
            </ul>
          </div>

          {/* 회계 안전성 */}
          <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4">
            <p className="text-[11px] font-bold text-emerald-800 mb-2 uppercase tracking-wider">🔒 회계·데이터 안전성</p>
            <ul className="text-xs text-emerald-800 space-y-1 leading-relaxed">
              <li>✅ 상담 데이터 hook (useConsultations / useConsultationDashboardStats) 무수정 — 호출 위치만 유지</li>
              <li>✅ 상태 전이 로직 (lib/consultation/transitions.ts) 무수정</li>
              <li>✅ 알림톡 흐름 (Make 웹훅) 무수정</li>
              <li>✅ Google Calendar 동기화 무수정</li>
              <li>⚠ lib/utils/format.ts의 `change_requested → orange-100` 한 줄만 시맨틱 색으로 정규화 (B안 채택 시)</li>
            </ul>
          </div>

          {/* 결정 안내 */}
          <div className="rounded-2xl border-2 border-stone-900 bg-stone-900 text-white p-5 text-center">
            <p className="text-sm font-bold mb-1">📌 사장님 결정 대기 중</p>
            <p className="text-xs opacity-80 leading-relaxed">
              <b>A안 / B안</b> 중 선택 → 클로드가 실제 <code className="bg-white/10 px-1.5 py-0.5 rounded">/consultations</code> + 하위 컴포넌트에 적용
              <br />
              <span className="opacity-60">"B안으로 가자" 또는 "B안 + ○○ 조정" 모두 가능. 적용 후 § 즉시 삭제 + push.</span>
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
