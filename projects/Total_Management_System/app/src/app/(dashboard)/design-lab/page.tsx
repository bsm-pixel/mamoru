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
import { InspectionForm } from '@/components/repairs/inspection-form';

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
            § 복원수리 수리내역서 — 디지털 핀 마킹 폼 (모바일 데모)
            실제 InspectionForm(demo 모드): 핀 마킹·멘트 프리셋 모두 동작.
            데모라 사진은 로컬 미리보기, 저장은 비활성(API/스토리지 호출 X).
            채택 결정 후 이 § 삭제.
            ═══════════════════════════════════════════════════════════ */}
        <section className="space-y-4">
          <div className="flex items-baseline gap-2 flex-wrap">
            <h2 className="text-lg font-bold text-stone-800">§ 수리내역서 — 핀 마킹 폼 (모바일 데모)</h2>
            <span className="text-[11px] text-stone-500">실제 동작 · 저장 안 됨</span>
          </div>
          <p className="text-xs text-stone-500 leading-relaxed">
            ① 사진 「촬영/업로드」로 가위 사진을 넣고 → ② 상처 유형 칩 선택 → ③ 사진을 탭해 핀을 찍어보세요(핀 탭 = 삭제).
            ④ 아래 「진단 멘트」 칩을 체크하면 줄바꿈으로 본문에 들어갑니다. 이대로 갈지 보시고 알려주세요.
            <br />
            <span className="text-stone-400">※ 데모 모드: 사진은 이 화면 안에서만 보이고 어디에도 저장되지 않습니다.</span>
          </p>

          {/* 모바일(390px) 프레임 */}
          <div className="flex justify-center bg-stone-100 rounded-2xl py-6">
            <div className="w-[390px] max-w-full bg-[#FAF9F7] rounded-[24px] shadow-lg border border-stone-200 overflow-hidden">
              <div className="p-3 max-h-[78vh] overflow-y-auto">
                <InspectionForm
                  demo
                  repairId="demo"
                  existingInspections={[]}
                  totalScissors={1}
                  initialComment=""
                />
              </div>
            </div>
          </div>
        </section>

        <div className="text-center text-[11px] text-stone-400 pt-4 border-t border-stone-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}
