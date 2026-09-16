'use client';

import { useState, type KeyboardEvent } from 'react';

/**
 * 검색 입력에서 ↑↓ 로 고르고 Enter 로 바로 선택 (2026-09-16)
 *
 * 사장님 요청: "검색해서 하단에 하나 뜨면 엔터로 바로 입력되는 기본적인 느낌"
 *   전엔 `sales/new` 한 곳만 이 동작이 있었고 나머지 선택 화면은 마우스로 클릭해야 했다.
 *   (모달 배경 버그와 같은 원인 — 공용 처리 없이 한 곳만 구현해서 퍼지지 않았다)
 *
 * 규칙
 *   ↑ ↓     : 활성 행 이동
 *   Enter   : ① 활성 행이 있으면 그것  ② 없고 결과가 **딱 1개**면 그것
 *             → 결과가 2개 이상인데 지정을 안 했으면 **아무 일도 안 한다**(오선택 방지)
 *   🔑 한글 조합 중(isComposing) Enter 는 무시 — 조합 확정용 Enter 가 선택으로 새면 안 된다
 *
 * 사용
 *   const pick = useEnterSelect(filtered, (p) => addProduct(p));
 *   <input onKeyDown={pick.onKeyDown} … />
 *   {filtered.map((p, i) => (
 *     <li key={p.id} className={pick.isActive(i) ? 'bg-neutral-100' : ''} onMouseEnter={() => pick.setActiveIdx(i)}>
 *   ))}
 */
export function useEnterSelect<T>(items: T[], onPick: (item: T) => void) {
  const [rawIdx, setActiveIdx] = useState(-1);
  // 검색어가 바뀌어 결과가 줄면 활성 행 무효화.
  // useEffect 로 되돌리면 한 박자 늦게(추가 렌더) 반영돼 그 사이 Enter 가 엉뚱한 행을 고를 수 있다 → 렌더 중 계산.
  const activeIdx = rawIdx < items.length ? rawIdx : -1;

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    // 한글 등 IME 조합 중에는 Enter 가 "조합 확정"이라 선택으로 쓰면 안 된다
    if ((e.nativeEvent as unknown as { isComposing?: boolean }).isComposing) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setActiveIdx((i) => Math.min(i + 1, items.length - 1));
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveIdx((i) => Math.max(i - 1, 0));
      return;
    }
    if (e.key !== 'Enter') return;

    if (activeIdx >= 0 && activeIdx < items.length) {
      e.preventDefault();
      onPick(items[activeIdx]);
      setActiveIdx(-1);
      return;
    }
    // 지정 없이 결과가 딱 1개일 때만 자동 선택 (2개 이상이면 무동작 — 오선택 방지)
    if (items.length === 1) {
      e.preventDefault();
      onPick(items[0]);
      setActiveIdx(-1);
    }
  };

  return {
    activeIdx,
    setActiveIdx,
    onKeyDown,
    isActive: (i: number) => i === activeIdx,
  };
}
