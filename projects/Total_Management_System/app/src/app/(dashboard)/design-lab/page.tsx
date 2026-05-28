'use client';

/**
 * /design-lab — TMS 디자인 모니터 (사장님 + 클로드 협업 도구)
 *
 * 2026-05-26 사장님 운영 룰:
 *   ▶ 진행 중인 디자인 작업만 표시 → 완료 후 § 자동 삭제 (회전 도구)
 *   ▶ 영구 박제는 운영 페이지 + memory/docs
 *
 * 현재 상태: 진행 중인 디자인 작업 없음 (클린 골격)
 *
 * 직전 완료:
 *   - 2026-05-27 § 대시보드 리모델 시안 B+ → /dashboard 적용
 *   - 2026-05-27 § 상담 페이지군 시안 B → /consultations 적용 (전체 탭 추가)
 *   - 2026-05-27 § 어두운 매출 카드 A2(그라데이션) → /repairs + /sales 동시 적용
 *
 * 운영 데이터 호출 X · 사이드바 메뉴 미노출 · URL 직접 접근만
 */

import { Topbar } from '@/components/layout/topbar';

export default function DesignLabPage() {
  return (
    <>
      <Topbar title="🎨 디자인 모니터" />
      <div className="px-6 py-6 space-y-8 max-w-[1100px] mx-auto bg-stone-50 min-h-screen">
        <div className="px-5 py-4 rounded-2xl bg-stone-900 text-white">
          <div className="flex items-center gap-2 mb-1">
            <span className="text-base font-bold">🎨 TMS 디자인 모니터</span>
            <span className="text-[10px] px-2 py-0.5 rounded bg-white/15 uppercase tracking-wider">internal tool</span>
          </div>
          <p className="text-xs opacity-80 leading-relaxed">
            사장님 + 클로드 협업 디자인 검토 도구. 진행 중인 작업만 표시하고 완료 후 비웁니다.
            <br />
            <span className="opacity-60">운영 데이터 X · 사이드바 메뉴 미노출 · URL 직접 접근만</span>
          </p>
        </div>

        {/* ═══════════════════════════════════════════════════════════
            § 고객 후기작성 페이지 (현재 형태) — 복원수리 / 상담 / 제품구매
            실제 page_review.html 을 iframe 으로 임베드 (uid=demo → urlName fallback 으로 폼 렌더).
            항상 실제 페이지와 동기화. 사장님 검토 후 수정 지시 → 클로드 반영 → 이 § 삭제.
            ═══════════════════════════════════════════════════════════ */}
        <section className="space-y-4">
          <div className="flex items-baseline gap-2 flex-wrap">
            <h2 className="text-lg font-bold text-stone-800">§ 고객 후기작성 페이지 — 현재 형태</h2>
            <span className="text-[11px] text-stone-500">알림톡 링크로 고객이 여는 실제 페이지 (3 유형)</span>
          </div>
          <p className="text-xs text-stone-500 leading-relaxed">
            아래는 <span className="font-semibold text-stone-700">page.mamoru.kr/projects/reviews/page_review.html</span> 실제 페이지를
            유형별로 임베드한 것입니다. 보면서 수정할 점을 말씀해 주세요.
            <br />
            <span className="text-stone-400">※ 데모 모드(uid=demo)라 이름은 &quot;홍**&quot;(마스킹)으로 표시되고, 제품구매는 제품선택 화면을 건너뛰고 후기 폼이 바로 보입니다(실제 사용 시 제품선택 화면 먼저 등장).</span>
          </p>

          <div className="grid gap-4 lg:grid-cols-3">
            {[
              { key: 'repair', label: '복원수리 후기', sub: 'restoration', tone: 'bg-amber-50 text-amber-700' },
              { key: 'consult', label: '상담 후기 (출장)', sub: 'field_request', tone: 'bg-violet-50 text-violet-700' },
              { key: 'purchase', label: '제품구매 후기', sub: '', tone: 'bg-emerald-50 text-emerald-700' },
            ].map((t) => {
              const src = `https://page.mamoru.kr/projects/reviews/page_review.html?type=${t.key}&uid=demo&name=${encodeURIComponent('홍길동')}${t.sub ? `&subtype=${t.sub}` : ''}`;
              return (
                <div key={t.key} className="rounded-2xl border border-stone-200 bg-white overflow-hidden flex flex-col">
                  <div className="flex items-center justify-between px-4 py-2.5 border-b border-stone-100">
                    <span className={`text-xs font-semibold px-2.5 py-1 rounded-full ${t.tone}`}>{t.label}</span>
                    <a href={src} target="_blank" rel="noopener noreferrer" className="text-[11px] text-stone-400 hover:text-stone-600 underline">새 탭</a>
                  </div>
                  <iframe
                    src={src}
                    title={`후기작성 — ${t.label}`}
                    className="w-full bg-[#FAF9F7]"
                    style={{ height: 940, border: 'none' }}
                    loading="lazy"
                  />
                </div>
              );
            })}
          </div>
        </section>

        <div className="text-center text-[11px] text-stone-400 pt-4 border-t border-stone-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}
