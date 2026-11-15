'use client';

import { Topbar } from '@/components/layout/topbar';
import { LabelEditor } from '@/components/labels/label-editor';

export default function LabelsPage() {
  return (
    <>
      <Topbar title="라벨 디자이너" />
      {/* 2026-11-15 모바일에서는 편집판이 옆으로 283px 넘쳤다 → PC 전용으로 안내 */}
      <div className="lg:hidden px-4 py-10 text-center">
        <p className="text-sm font-bold text-neutral-800">라벨 디자이너는 PC에서 써 주세요</p>
        <p className="mt-1.5 text-xs text-neutral-500">라벨 요소를 끌어서 배치하는 화면이라 넓은 화면이 필요합니다.</p>
      </div>
      <div className="hidden lg:block px-4 md:px-6 py-5">
        <p className="text-xs text-neutral-500 mb-4">라벨 요소를 직접 드래그·조정해 배치하고 저장하면, 제품·시리얼 라벨 출력에 바로 반영됩니다.</p>
        <LabelEditor />
      </div>
    </>
  );
}
