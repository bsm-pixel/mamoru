'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { useHotkeys } from '@/hooks/use-hotkeys';

/** 단축키 1줄 표시용 */
function Row({ keys, desc }: { keys: string[]; desc: string }) {
  return (
    <div className="flex items-center justify-between gap-3 py-1.5">
      <span className="text-sm text-neutral-600">{desc}</span>
      <span className="flex items-center gap-1 shrink-0">
        {keys.map((k) => (
          <kbd
            key={k}
            className="px-1.5 py-0.5 text-[11px] font-sans font-semibold text-neutral-600 bg-neutral-50 border border-neutral-300 rounded shadow-[0_1px_0_rgba(0,0,0,0.06)]"
          >
            {k}
          </kbd>
        ))}
      </span>
    </div>
  );
}

/**
 * 전역 단축키 도움말. `?`(Shift+/) 로 열고 Esc 로 닫힘(Modal 처리).
 * 대시보드 레이아웃에 한 번만 마운트 → 모든 화면에서 사용 가능.
 */
export function ShortcutHelp() {
  const [open, setOpen] = useState(false);
  useHotkeys([{ combo: '?', handler: () => setOpen((o) => !o) }]);

  return (
    <Modal open={open} onClose={() => setOpen(false)} title="⌨️ 단축키" className="max-w-md">
      <div className="divide-y divide-neutral-100">
        <div className="pb-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">공통</p>
          <Row keys={['?']} desc="이 도움말 열기 / 닫기" />
          <Row keys={['Esc']} desc="열린 창·상세 닫기 / 검색어 지우기" />
          <Row keys={['Ctrl', 'K']} desc="빠른 찾기 — 고객 · 판매번호 · 제품 · 메뉴" />
          <Row keys={['/']} desc="검색창으로 이동" />
          <Row keys={['Ctrl', '[']} desc="이전 화면으로 (뒤로)" />
          <Row keys={['Ctrl', ']']} desc="다음 화면으로 (앞으로)" />
        </div>
        <div className="py-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">목록 화면</p>
          <Row keys={['↑', '↓']} desc="줄 이동 (검색 칸에서 ↓ = 첫 줄)" />
          <Row keys={['Enter']} desc="그 건의 「다음 할 일」 버튼으로" />
          <Row keys={['N']} desc="새로 만들기" />
          <Row keys={['1', '~', '9']} desc="상태 탭 바꾸기" />
        </div>
        <div className="py-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">어느 화면에서나</p>
          <Row keys={['F2']} desc="판매 입력 열기" />
          <Row keys={['F4']} desc="거래처 매출 열기" />
        </div>
        <div className="py-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">판매 입력</p>
          <Row keys={['Ctrl', 'Enter']} desc="판매 등록" />
          <Row keys={['F7']} desc="제품 검색 칸으로 (바코드 찍기)" />
          <Row keys={['F8']} desc="고객 칸으로" />
          <Row keys={['F9']} desc="결제 수단 바꾸기" />
          <Row keys={['←', '→']} desc="선택 버튼 묶음에서 옆으로 고르기" />
          <Row keys={['Tab']} desc="고객 → 제품 → 결제 수단 → 메모 → 판매 등록" />
        </div>
        <div className="py-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">입력 창</p>
          <Row keys={['Ctrl', 'Enter']} desc="저장" />
          <Row keys={['Enter']} desc="다음 칸으로 (한 줄 칸)" />
          <Row keys={['Esc']} desc="닫기 — 넣은 내용이 있으면 한 번 묻기" />
        </div>
        <div className="pt-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">입력 폼 · 제품 검색</p>
          <Row keys={['Ctrl', 'S']} desc="저장 / 등록" />
          <Row keys={['↑', '↓']} desc="검색 결과 행 이동(지정)" />
          <Row keys={['Enter']} desc="지정한 행(또는 결과 1개) 담기 · 정확한 코드는 바로 담기" />
        </div>
      </div>
      <p className="mt-4 text-[11px] text-neutral-400">
        입력창에 타이핑 중일 때는 문자 단축키(<kbd className="px-1 border border-neutral-200 rounded">/</kbd>,{' '}
        <kbd className="px-1 border border-neutral-200 rounded">?</kbd>)가 동작하지 않습니다.
      </p>
    </Modal>
  );
}
