'use client';

import { Topbar } from '@/components/layout/topbar';
import { LabelEditor } from '@/components/labels/label-editor';

export default function LabelsPage() {
  return (
    <>
      <Topbar title="라벨 디자이너" />
      <div className="px-4 md:px-6 py-5">
        <p className="text-xs text-neutral-500 mb-4">라벨 요소를 직접 드래그·조정해 배치하고 저장하면, 제품·시리얼 라벨 출력에 바로 반영됩니다.</p>
        <LabelEditor />
      </div>
    </>
  );
}
