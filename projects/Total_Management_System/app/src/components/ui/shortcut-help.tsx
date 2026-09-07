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
          <Row keys={['/']} desc="검색창으로 이동" />
          <Row keys={['Ctrl', '[']} desc="이전 화면으로 (뒤로)" />
          <Row keys={['Ctrl', ']']} desc="다음 화면으로 (앞으로)" />
        </div>
        <div className="py-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">판매 관리</p>
          <Row keys={['F2']} desc="판매 입력 열기" />
          <Row keys={['F4']} desc="거래처 매출 열기" />
        </div>
        <div className="pt-2">
          <p className="text-[11px] font-semibold text-neutral-400 uppercase tracking-wide mb-1">입력 폼 · 제품 검색</p>
          <Row keys={['Ctrl', 'S']} desc="저장 / 등록" />
          <Row keys={['↑', '↓']} desc="검색 결과 행 이동(지정)" />
          <Row keys={['Enter']} desc="지정한 행(또는 결과 1개) 담기" />
        </div>
      </div>
      <p className="mt-4 text-[11px] text-neutral-400">
        입력창에 타이핑 중일 때는 문자 단축키(<kbd className="px-1 border border-neutral-200 rounded">/</kbd>,{' '}
        <kbd className="px-1 border border-neutral-200 rounded">?</kbd>)가 동작하지 않습니다.
      </p>
    </Modal>
  );
}
