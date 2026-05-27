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

        <section className="border-2 border-dashed border-stone-300 rounded-2xl p-12 text-center bg-white">
          <div className="text-4xl mb-3">🎨</div>
          <h3 className="text-base font-bold text-stone-700 mb-2">진행 중인 디자인 작업 없음</h3>
          <p className="text-xs text-stone-500 max-w-md mx-auto leading-relaxed">
            새 페이지 디자인 검토가 필요하면 클로드에게 한 마디 하세요.
            <br />
            <span className="font-semibold text-stone-700">&quot;디자인 모니터에 § XXX 추가해줘&quot;</span>
            <br />
            → 비교 옵션 1~3개가 여기에 렌더됩니다.
          </p>
          <div className="mt-5 text-[11px] text-stone-400 space-y-1 text-left max-w-sm mx-auto">
            <p>운영 룰:</p>
            <p>• 매출 카드 / 목록 카드 / 상세 패널 / 헤더 등 모든 UI 영역 가능</p>
            <p>• 비교 후 결정 → 클로드가 실제 페이지에 적용</p>
            <p>• 적용 완료 후 § 자동 삭제 → 페이지 클린 상태 유지</p>
            <p>• 채택안 박제는 memory/docs 에 (영구), 여기는 회전 도구</p>
          </div>
        </section>

        <div className="text-center text-[11px] text-stone-400 pt-4 border-t border-stone-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}
