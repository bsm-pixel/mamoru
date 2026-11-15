import type { KeyboardEvent } from 'react';

/**
 * 선택 버튼 묶음(매장/출장/톡, 카드/현금/이체 …)을 ← → 로 고르기 (2026-11-15)
 *
 * 묶음을 감싼 요소에 onKeyDown 으로 단다. 방향키를 누르면 옆 버튼으로 초점이 가고 바로 선택된다.
 * 고른 버튼만 tabIndex 를 주고 나머지는 -1 로 두면, Tab 한 번에 묶음 전체를 지나간다.
 *
 *   <div role="radiogroup" onKeyDown={groupArrowKeys}> <button tabIndex={on ? 0 : -1} …/> … </div>
 */
export function groupArrowKeys(e: KeyboardEvent<HTMLElement>) {
  if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
  if (e.ctrlKey || e.metaKey || e.altKey) return;
  const btns = Array.from(e.currentTarget.querySelectorAll<HTMLButtonElement>('button:not(:disabled)'));
  const i = btns.indexOf(document.activeElement as HTMLButtonElement);
  if (i < 0) return;
  e.preventDefault();
  const next = btns[(i + (e.key === 'ArrowRight' ? 1 : btns.length - 1)) % btns.length];
  next.focus();
  next.click();
}
