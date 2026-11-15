'use client';

import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { Search, X } from 'lucide-react';
import { matchSearch } from '@/lib/utils/search-match';

/**
 * 검색해서 고르는 칸 — 한 부품 (2026-11-15)
 *
 * 전엔 화면마다 제각각이었다(재고 조정은 검색 없는 165줄 목록, 시리얼은 제품 전체가 버튼으로 깔림 …).
 * 규칙은 하나:
 *   글자 입력 즉시 걸러짐 / ↓↑ 이동 / Enter = 강조된 것(없으면 첫 번째) / Esc = 목록만 닫기
 *   Tab = 고르고 다음 칸 / 결과가 없고 onCreate 가 있으면 Enter = 새로 등록
 * 검색은 lib/utils/search-match.ts — 이름 · 코드 · 전화 뒤 4자리 · 초성, 띄어쓰기 · 하이픈 무시.
 *
 *   <SearchPick items={products} value={picked} onPick={setPicked} onClear={() => setPicked(null)}
 *     getKey={(p) => p.id} getLabel={(p) => p.name} getSub={(p) => p.sku} getSearch={(p) => [p.name, p.sku]} />
 */
export interface SearchPickProps<T> {
  items: T[];
  value: T | null;
  onPick: (item: T) => void;
  onClear?: () => void;
  getKey: (item: T) => string;
  getLabel: (item: T) => string;
  /** 이름 아래 작은 글씨(코드 · 전화 등) */
  getSub?: (item: T) => string | null | undefined;
  /** 줄 오른쪽 끝(재고 · 금액 등) */
  getRight?: (item: T) => ReactNode;
  /** 검색에 쓸 글자들. 없으면 이름 + 작은 글씨 */
  getSearch?: (item: T) => Array<string | null | undefined>;
  placeholder?: string;
  /** 결과가 없을 때 Enter 로 새로 만들기 */
  onCreate?: (query: string) => void;
  createLabel?: string;
  /** 고른 뒤에도 칸을 비우고 계속 고르기(여러 개 담기용) */
  keepOpen?: boolean;
  autoFocus?: boolean;
  id?: string;
  tabIndex?: number;
  /** 한 번에 보여 줄 최대 줄 수 (기본 50) */
  limit?: number;
  disabled?: boolean;
  className?: string;
}

export function SearchPick<T>({
  items, value, onPick, onClear, getKey, getLabel, getSub, getRight, getSearch,
  placeholder = '이름 · 코드로 검색', onCreate, createLabel = '새로 등록', keepOpen,
  autoFocus, id, tabIndex, limit = 50, disabled, className = '',
}: SearchPickProps<T>) {
  const [query, setQuery] = useState('');
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(-1);
  const wrapRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const activeRef = useRef<HTMLLIElement>(null);

  const results = useMemo(() => {
    const out: T[] = [];
    for (const it of items) {
      const texts = getSearch ? getSearch(it) : [getLabel(it), getSub?.(it)];
      if (matchSearch(query, ...texts)) { out.push(it); if (out.length >= limit) break; }
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items, query, limit]);

  useEffect(() => { activeRef.current?.scrollIntoView({ block: 'nearest' }); }, [active]);
  useEffect(() => {
    if (autoFocus && !(typeof window !== 'undefined' && window.matchMedia?.('(pointer: coarse)').matches)) inputRef.current?.focus();
  }, [autoFocus]);
  // 바깥을 누르면 목록 닫기
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => { if (!wrapRef.current?.contains(e.target as Node)) setOpen(false); };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [open]);

  const pick = (it: T) => {
    onPick(it);
    setQuery('');
    setActive(-1);
    setOpen(!!keepOpen);
    if (keepOpen) inputRef.current?.focus();
  };

  // 고른 것이 있으면 칸 대신 고른 값 + 지우기
  if (value && !keepOpen) {
    return (
      <div className={`flex items-center gap-2 h-9 px-3 rounded-lg border border-neutral-200 bg-white ${className}`}>
        <span className="flex-1 min-w-0 text-sm font-semibold text-stone-900 truncate">{getLabel(value)}</span>
        {getSub?.(value) && <span className="text-xs text-neutral-400 shrink-0 truncate max-w-[40%]">{getSub(value)}</span>}
        {onClear && !disabled && (
          <button type="button" title="선택 지우기" aria-label="선택 지우기" onClick={() => { onClear(); setTimeout(() => inputRef.current?.focus(), 0); }}
            className="w-5 h-5 rounded-full bg-neutral-200 flex items-center justify-center hover:bg-neutral-300 shrink-0">
            <X size={10} />
          </button>
        )}
      </div>
    );
  }

  const showCreate = !!onCreate && query.trim().length > 0 && results.length === 0;

  return (
    <div ref={wrapRef} className={`relative ${className}`}>
      <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400 pointer-events-none" />
      <input
        ref={inputRef}
        id={id}
        tabIndex={tabIndex}
        type="text"
        role="combobox"
        aria-expanded={open}
        aria-autocomplete="list"
        autoComplete="off"
        disabled={disabled}
        value={query}
        onChange={(e) => { setQuery(e.target.value); setActive(-1); setOpen(true); }}
        onFocus={() => setOpen(true)}
        onKeyDown={(e) => {
          if (e.nativeEvent.isComposing) return; // 한글 조합 중의 Enter 는 글자 확정용
          if (e.key === 'ArrowDown') { e.preventDefault(); setOpen(true); setActive((i) => Math.min(i + 1, results.length - 1)); return; }
          if (e.key === 'ArrowUp') { e.preventDefault(); setActive((i) => Math.max(i - 1, 0)); return; }
          if (e.key === 'Escape') {
            if (open) { e.preventDefault(); e.stopPropagation(); e.nativeEvent.stopImmediatePropagation(); setOpen(false); setActive(-1); }
            return;
          }
          if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey) {
            if (!open && !query) return;
            e.preventDefault();
            e.stopPropagation();
            const target = results[active >= 0 ? active : 0];
            if (target) pick(target);
            else if (showCreate) { onCreate!(query.trim()); setOpen(false); }
            return;
          }
          if (e.key === 'Tab' && !e.shiftKey && open && query.trim()) {
            // 고르고 다음 칸으로(기본 Tab 이동은 그대로 일어난다)
            const target = results[active >= 0 ? active : 0];
            if (target) pick(target);
          }
        }}
        placeholder={placeholder}
        className="w-full h-9 pl-8 pr-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300 disabled:opacity-50"
      />
      {open && !disabled && (
        <ul role="listbox" className="absolute z-50 left-0 right-0 mt-1 max-h-64 overflow-y-auto rounded-lg border border-neutral-200 bg-white">
          {results.map((it, i) => (
            <li
              key={getKey(it)}
              ref={i === active ? activeRef : undefined}
              role="option"
              aria-selected={i === active}
              onMouseDown={(e) => { e.preventDefault(); pick(it); }}
              onMouseEnter={() => setActive(i)}
              className={`flex items-center justify-between gap-3 px-3 py-2 cursor-pointer border-b border-neutral-50 last:border-b-0 ${i === active ? 'bg-stone-100' : 'hover:bg-stone-50'}`}
            >
              <span className="min-w-0">
                <span className="block text-sm font-medium text-stone-900 truncate">{getLabel(it)}</span>
                {getSub?.(it) && <span className="block text-[11px] text-neutral-400 truncate">{getSub(it)}</span>}
              </span>
              {getRight && <span className="text-xs text-neutral-500 tabular-nums shrink-0">{getRight(it)}</span>}
            </li>
          ))}
          {results.length === 0 && !showCreate && (
            <li className="px-3 py-3 text-xs text-neutral-400 text-center">{query ? '맞는 것이 없습니다' : '목록이 비어 있습니다'}</li>
          )}
          {showCreate && (
            <li onMouseDown={(e) => { e.preventDefault(); onCreate!(query.trim()); setOpen(false); }}
              className="px-3 py-2.5 text-sm font-semibold text-stone-900 cursor-pointer bg-stone-50 hover:bg-stone-100">
              「{query.trim()}」 {createLabel} <span className="text-[11px] font-normal text-neutral-400 ml-1">Enter</span>
            </li>
          )}
          {results.length >= limit && (
            <li className="px-3 py-1.5 text-[11px] text-neutral-400 text-center">더 있습니다 — 글자를 더 치면 좁혀집니다</li>
          )}
        </ul>
      )}
    </div>
  );
}
