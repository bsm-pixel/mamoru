'use client';

import { useState } from 'react';
import { ChevronDown, ChevronUp, Sparkles } from 'lucide-react';
import { useDemoPO } from './use-demo-po';
import { POForm } from './POForm';
import { LabelPreview } from './LabelPreview';
import { InboundMatchMobile } from './InboundMatchMobile';
import { InboundGridPC } from './InboundGridPC';
import { PromoteModal } from './PromoteModal';
import { MobileFrame } from './MobileFrame';

/**
 * § 1688 중국 사입 — 디자인 모니터 시안
 *
 * Phase A: useState 로컬 상태만 사용. API/DB 호출 없음.
 * 사장님 피드백 받아가며 _sections/sourcing-1688/ 안에서만 수정.
 * OK 사인 → Phase B (운영 매입관리 적용) + § 즉시 삭제.
 */
export function Sourcing1688Section() {
  const api = useDemoPO();
  const [promoteItemId, setPromoteItemId] = useState<string | null>(null);

  // STEP 토글 (사장님이 한 단계씩 집중해서 보고 싶을 때)
  const [openSteps, setOpenSteps] = useState({
    s1: true,
    s2: true,
    s3: true,
    s4: true,
  });
  const toggle = (key: keyof typeof openSteps) =>
    setOpenSteps((cur) => ({ ...cur, [key]: !cur[key] }));

  return (
    <section className="space-y-6">
      {/* § 헤더 */}
      <div className="rounded-2xl border-2 border-amber-300 bg-amber-50 p-5">
        <div className="flex items-center gap-2 mb-2">
          <span className="text-base font-bold text-stone-900">§ 1688 중국 사입</span>
          <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-amber-200 text-amber-900 uppercase tracking-wider">
            데모 모드
          </span>
        </div>
        <p className="text-xs text-stone-700 leading-relaxed">
          PO 작성 → 라벨 인쇄(라벨프린터 1클릭) → QR 스캔 입고매칭(폰) → PC 그리드 현황 → 정식 SKU 승격 흐름.
          <br />
          <span className="text-stone-500">
            ※ 모든 상태는 화면 새로고침 시 초기화. 저장/업로드/API 호출 없음. 운영 데이터 무영향.
          </span>
        </p>
        <div className="mt-3 flex items-center gap-2 text-[11px] text-stone-600">
          <Sparkles size={12} className="text-amber-600" />
          <span>
            처음 보시면 <strong>STEP 1 우측 상단 [예시 채우기]</strong> 버튼을 눌러 흐름을 따라가 보세요.
          </span>
        </div>
      </div>

      {/* STEP 1. PO 작성 */}
      <StepBlock
        n={1}
        title="PO 작성"
        sub="회사명/품목명/단가/MOQ/특징/1688 URL 입력 → 자동 스티커번호 부여"
        open={openSteps.s1}
        onToggle={() => toggle('s1')}
      >
        <POForm api={api} />
      </StepBlock>

      {/* STEP 2. 라벨 미리보기 */}
      <StepBlock
        n={2}
        title="라벨 미리보기 · 인쇄"
        sub="사이즈 선택 → STEP 1 입력 기반 자동 생성 · QR 스캔 시 해당 품목 입고매칭 페이지 직진입"
        open={openSteps.s2}
        onToggle={() => toggle('s2')}
      >
        <LabelPreview po={api.po} />
      </StepBlock>

      {/* STEP 3+4. 입고 매칭 (모바일 + PC 동시 보기) */}
      <StepBlock
        n={3}
        title="입고 매칭 — 모바일(폰 QR) ↔ PC 그리드 동시"
        sub="좌측 모바일 화면에서 매칭하면 우측 PC 그리드가 실시간 emerald 전환"
        open={openSteps.s3}
        onToggle={() => toggle('s3')}
      >
        <div className="flex flex-col xl:flex-row gap-6 items-start">
          <MobileFrame title="📱 폰에서 QR 스캔 후 진입한 화면">
            <InboundMatchMobile api={api} />
          </MobileFrame>
          <div className="flex-1 min-w-0 w-full">
            <div className="text-[11px] text-stone-500 mb-2 font-medium">
              💻 PC 매입관리 → 1688 PO → 그리드 매칭 뷰
            </div>
            <InboundGridPC api={api} onPromote={(id) => setPromoteItemId(id)} />
          </div>
        </div>
      </StepBlock>

      {/* STEP 4. 정식 채택 결과 */}
      <StepBlock
        n={4}
        title="정식 SKU 승격 결과"
        sub="채택된 품목은 products 테이블에 INSERT (운영 시) → 재고/판매 흐름 진입"
        open={openSteps.s4}
        onToggle={() => toggle('s4')}
      >
        <PromotedList api={api} />
      </StepBlock>

      {/* 모달 */}
      <PromoteModal
        api={api}
        itemId={promoteItemId}
        onClose={() => setPromoteItemId(null)}
      />
    </section>
  );
}

function StepBlock({
  n,
  title,
  sub,
  open,
  onToggle,
  children,
}: {
  n: number;
  title: string;
  sub: string;
  open: boolean;
  onToggle: () => void;
  children: React.ReactNode;
}) {
  return (
    <div className="rounded-2xl bg-white border border-stone-200 overflow-hidden">
      <button
        type="button"
        onClick={onToggle}
        className="w-full px-5 py-4 flex items-center gap-3 hover:bg-stone-50 text-left"
      >
        <span className="inline-flex items-center justify-center w-9 h-9 rounded-full bg-stone-900 text-white text-sm font-bold flex-shrink-0">
          {n}
        </span>
        <div className="flex-1 min-w-0">
          <div className="font-bold text-stone-900">{title}</div>
          <div className="text-xs text-stone-500 mt-0.5 truncate">{sub}</div>
        </div>
        {open ? (
          <ChevronUp size={18} className="text-stone-400 flex-shrink-0" />
        ) : (
          <ChevronDown size={18} className="text-stone-400 flex-shrink-0" />
        )}
      </button>
      {open && <div className="px-5 pb-5 border-t border-stone-100 pt-4">{children}</div>}
    </div>
  );
}

function PromotedList({ api }: { api: ReturnType<typeof useDemoPO> }) {
  const promoted = api.po.items.filter((it) => it.inspection_status === 'promoted');
  const rejected = api.po.items.filter((it) => it.inspection_status === 'rejected');

  if (promoted.length === 0 && rejected.length === 0) {
    return (
      <div className="text-center py-8 text-xs text-stone-500">
        아직 정식 채택/보류된 품목이 없습니다.
        <br />
        STEP 3 그리드에서 매칭된 셀에 마우스 올리면 [정식 채택]·[보류] 액션이 나타납니다.
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {promoted.length > 0 && (
        <div>
          <div className="text-[11px] font-bold text-stone-500 uppercase tracking-wider mb-2">
            정식 채택 {promoted.length}건
          </div>
          <div className="space-y-2">
            {promoted.map((it) => (
              <div
                key={it.id}
                className="flex items-center gap-3 p-3 rounded-lg bg-stone-900 text-white"
              >
                <span className="font-mono text-xs bg-white/15 px-2 py-1 rounded font-bold">
                  {it.promoted_sku}
                </span>
                <div className="flex-1 text-sm font-medium truncate">{it.promoted_name}</div>
                <span className="text-[11px] opacity-70">
                  ←사입 {it.sticker_no.split('-').pop()} · {it.quantity}개
                </span>
              </div>
            ))}
          </div>
        </div>
      )}
      {rejected.length > 0 && (
        <div>
          <div className="text-[11px] font-bold text-stone-500 uppercase tracking-wider mb-2">
            보류 {rejected.length}건
          </div>
          <div className="space-y-2">
            {rejected.map((it) => (
              <div
                key={it.id}
                className="flex items-center gap-3 p-3 rounded-lg bg-rose-50 border border-rose-200"
              >
                <span className="font-mono text-xs bg-white px-2 py-1 rounded text-rose-700 font-bold border border-rose-200">
                  {it.sticker_no.split('-').pop()}
                </span>
                <div className="flex-1 text-sm font-medium text-stone-800 truncate">
                  {it.product_name}
                </div>
                <span className="text-[11px] text-stone-500">정식 SKU 미부여</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
