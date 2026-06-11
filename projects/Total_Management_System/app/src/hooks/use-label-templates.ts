'use client';

/**
 * 라벨 템플릿 저장/로드 — 코드 기본값(LABEL_TEMPLATES) 위에 사장님이 편집한 값을 settings에 덮어쓰기.
 * 편집기에서 저장 → `label.templates` 키(JSON) → 출력 시 이 효과값 사용.
 */

import { useMemo } from 'react';
import { useSetting, useUpdateSettings } from './use-settings';
import { LABEL_TEMPLATES, type LabelTemplate } from '@/lib/label/templates';

const KEY = 'label.templates';

/** 저장본 병합한 전체 템플릿 맵 */
export function useLabelTemplates(): Record<string, LabelTemplate> {
  const saved = useSetting<Record<string, LabelTemplate>>(KEY, {});
  return { ...LABEL_TEMPLATES, ...(saved || {}) };
}

/** id의 효과 템플릿(저장본 우선, 없으면 코드 기본값).
 *  ⚠️ useSetting이 매 렌더 새 객체를 반환 → 내용(직렬화) 기준 메모이즈해 안정 identity 보장(무한루프/불필요 재렌더 차단). */
export function useLabelTemplate(id: string): LabelTemplate {
  const saved = useSetting<Record<string, LabelTemplate>>(KEY, {});
  const savedJson = JSON.stringify(saved?.[id] ?? null);
  return useMemo(() => (saved?.[id] || LABEL_TEMPLATES[id]), [savedJson, id]); // eslint-disable-line react-hooks/exhaustive-deps
}

/** 편집 저장/복원 */
export function useSaveLabelTemplate() {
  const saved = useSetting<Record<string, LabelTemplate>>(KEY, {});
  const update = useUpdateSettings();
  return {
    save: (tpl: LabelTemplate) =>
      update.mutate([{ key: KEY, value: { ...(saved || {}), [tpl.id]: tpl } }]),
    reset: (id: string) => {
      const next = { ...(saved || {}) };
      delete next[id];
      update.mutate([{ key: KEY, value: next }]);
    },
    isPending: update.isPending,
  };
}
