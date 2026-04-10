'use client';

import { X } from 'lucide-react';

interface TagSelectorProps {
  availableTags: string[];
  selectedTags: string[];
  onChange: (tags: string[]) => void;
  readonly?: boolean;
}

/** 고객 태그 멀티셀렉트 칩 컴포넌트 */
export function TagSelector({ availableTags, selectedTags, onChange, readonly }: TagSelectorProps) {
  const unselected = availableTags.filter((t) => !selectedTags.includes(t));

  function toggle(tag: string) {
    if (readonly) return;
    if (selectedTags.includes(tag)) {
      onChange(selectedTags.filter((t) => t !== tag));
    } else {
      onChange([...selectedTags, tag]);
    }
  }

  if (availableTags.length === 0) return null;

  return (
    <div className="flex flex-wrap gap-1.5">
      {selectedTags.map((t) => (
        <button
          key={t}
          type="button"
          onClick={() => toggle(t)}
          disabled={readonly}
          className="flex items-center gap-1 px-2 py-0.5 text-xs bg-blue-50 text-blue-700 rounded-full transition hover:bg-blue-100 disabled:opacity-70 disabled:cursor-default"
        >
          {t}
          {!readonly && <X size={10} className="shrink-0" />}
        </button>
      ))}
      {!readonly && unselected.map((t) => (
        <button
          key={t}
          type="button"
          onClick={() => toggle(t)}
          className="px-2 py-0.5 text-xs bg-neutral-100 text-neutral-500 rounded-full transition hover:bg-neutral-200"
        >
          + {t}
        </button>
      ))}
    </div>
  );
}

/** 태그 읽기 전용 표시 (목록/상세에서 사용) */
export function TagBadges({ tags }: { tags?: string[] | null }) {
  if (!tags || tags.length === 0) return null;
  return (
    <div className="flex flex-wrap gap-1">
      {tags.map((t) => (
        <span key={t} className="px-1.5 py-0.5 text-[10px] bg-blue-50 text-blue-600 rounded-full">
          {t}
        </span>
      ))}
    </div>
  );
}
