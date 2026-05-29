'use client';

/**
 * /design-lab — TMS 디자인 모니터 (사장님 + 클로드 협업 도구)
 *
 * 2026-05-26 사장님 운영 룰:
 *   ▶ 진행 중인 디자인/기능 작업만 표시 → 완료 후 § 자동 삭제 (회전 도구)
 *   ▶ 영구 박제는 운영 페이지 + memory/docs
 *
 * 현재 진행: 복원수리 수리내역서 — 디지털 핀 마킹 폼 (모바일 데모, 채택 여부 확인용)
 *
 * 운영 데이터 호출 X · 사이드바 메뉴 미노출 · URL 직접 접근만
 */

import { Topbar } from '@/components/layout/topbar';
import { RepairReportDemo } from '@/components/repairs/repair-report-demo';

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
            사장님 + 클로드 협업 검토 도구. 진행 중인 작업만 표시하고 완료 후 비웁니다.
            <br />
            <span className="opacity-60">운영 데이터 X · 사이드바 메뉴 미노출 · URL 직접 접근만</span>
          </p>
        </div>

        {/* ═══════════════════════════════════════════════════════════
            § 수리내역서 — 좌(검수 입력) / 우(실시간 고객 화면) 데모
            좌측에서 마킹·멘트를 바꾸면 우측 고객 수리내역서가 즉시 반영.
            데모라 저장/업로드 없음. 채택 후 라이브 승격 → 이 § 삭제.
            ═══════════════════════════════════════════════════════════ */}
        <section className="space-y-4">
          <div className="flex items-baseline gap-2 flex-wrap">
            <h2 className="text-lg font-bold text-stone-800">§ 수리내역서 — 입력 ↔ 고객 화면 실시간</h2>
            <span className="text-[11px] text-stone-500">저장 안 됨 · 미리보기</span>
          </div>
          <p className="text-xs text-stone-500 leading-relaxed">
            좌측에서 ① 사진 넣고 → ② 상처 유형 선택 → ③ 사진 탭=점 / 드래그=범위(선) → ④ 멘트 칩 체크.
            바꾸는 즉시 <span className="font-semibold text-stone-700">우측 「고객 화면」</span>에 그대로 반영됩니다. 이대로 갈지 보시고 알려주세요.
          </p>

          <RepairReportDemo />
        </section>

        <div className="text-center text-[11px] text-stone-400 pt-4 border-t border-stone-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}
