'use client';

/**
 * 라벨 템플릿 저장/로드 — 코드 기본값(LABEL_TEMPLATES) 위에 사장님 편집/생성 템플릿을 settings에 덮어쓰기.
 * 편집기 저장 → `label.templates`(JSON) → 출력 시 이 효과값 사용.
 */

import { useMemo } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { useSetting, useUpdateSettings } from './use-settings';
import { LABEL_TEMPLATES, type LabelTemplate } from '@/lib/label/templates';

const KEY = 'label.templates';

/** 저장본 병합한 전체 템플릿 맵 */
export function useLabelTemplates(): Record<string, LabelTemplate> {
  const saved = useSetting<Record<string, LabelTemplate>>(KEY, {});
  return { ...LABEL_TEMPLATES, ...(saved || {}) };
}

/** id의 효과 템플릿(저장본 우선, 없으면 코드 기본값, 그래도 없으면 빈 안전 템플릿 — undefined 크래시 방지).
 *  ⚠️ useSetting이 매 렌더 새 객체 반환 → 내용(직렬화) 기준 메모이즈해 안정 identity(무한루프 방지). */
export function useLabelTemplate(id: string): LabelTemplate {
  const saved = useSetting<Record<string, LabelTemplate>>(KEY, {});
  const found = saved?.[id] || LABEL_TEMPLATES[id];
  const foundJson = JSON.stringify(found ?? null);
  return useMemo(
    () => found || { id, name: id, widthMm: 40, heightMm: 20, elements: [] },
    [foundJson, id], // eslint-disable-line react-hooks/exhaustive-deps
  );
}

/** 편집 저장/복원 — 낙관적 캐시 갱신(새 템플릿 즉시 사용, 흰화면 방지) */
export function useSaveLabelTemplate() {
  const saved = useSetting<Record<string, LabelTemplate>>(KEY, {});
  const update = useUpdateSettings();
  const qc = useQueryClient();

  const writeCache = (next: Record<string, LabelTemplate>) => {
    qc.setQueryData(['settings'], (old: Record<string, unknown> | undefined) => ({ ...(old || {}), [KEY]: next }));
  };

  return {
    save: (tpl: LabelTemplate) => {
      const next = { ...(saved || {}), [tpl.id]: tpl };
      writeCache(next); // 즉시 반영 → 새 id 바로 조회 가능
      update.mutate([{ key: KEY, value: next }]);
    },
    reset: (id: string) => {
      const next = { ...(saved || {}) };
      delete next[id];
      writeCache(next);
      update.mutate([{ key: KEY, value: next }]);
    },
    isPending: update.isPending,
  };
}
