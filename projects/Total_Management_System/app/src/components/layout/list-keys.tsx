'use client';

import { useEffect } from 'react';
import { isDialogOpen } from '@/hooks/use-hotkeys';

/**
 * 목록 화면 키보드 (2026-11-15) — 대시보드 레이아웃에 한 번만 둔다
 *
 *   /        검색 칸으로
 *   ↓ ↑      줄 이동 (옆 상세가 따라 바뀐다). 검색 칸에서 ↓ = 목록 첫 줄
 *   Enter    그 건의 「다음 할 일」 버튼으로 초점 (한 번 더 Enter 면 실행)
 *   Esc      상세 닫기 (화면이 useListEscape 로 받는다)
 *   N        새로 만들기
 *   1 ~ 9    상태 탭
 *
 * 화면은 표시만 달아 두면 된다(코드는 여기 한 곳):
 *   data-list-row (+ data-selected="true")   목록의 한 줄        ← 공용 표(DataGrid)는 이미 달려 있다
 *   data-list-search                          검색 칸             ← 공용 SearchInput 은 이미 달려 있다
 *   data-status-tabs                          상태 탭 묶음(안의 button 순서 = 1~9)
 *   data-new-action                           [새로 만들기] 버튼
 *   data-next-action                          「다음 할 일」 묶음(안의 첫 button)
 *
 * 글자를 치는 칸 안에서는 한 글자 단축키가 동작하지 않는다. 창이 떠 있을 때도 동작하지 않는다.
 * 브라우저가 쓰는 키(Ctrl+N · Ctrl+T · Ctrl+W · F5)는 건드리지 않는다.
 */

const visible = (el: Element) => {
  const r = (el as HTMLElement).getBoundingClientRect();
  return r.width > 0 && r.height > 0;
};
const all = (sel: string) => Array.from(document.querySelectorAll<HTMLElement>(sel)).filter(visible);

function isTyping(t: EventTarget | null): boolean {
  const el = t as HTMLElement | null;
  if (!el || !el.tagName) return false;
  return el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' || el.isContentEditable;
}

const DARK = /\bbg-(stone|neutral|zinc|gray|slate)-(800|900|950)\b|\bbg-black\b|\bbg-terracotta\b/;

/** 「다음 할 일」 버튼 찾기: 표시가 있으면 그것, 없으면 고른 줄보다 오른쪽(상세 칸)에 있는 첫 진한 버튼 */
function findNextAction(row: HTMLElement | null): HTMLButtonElement | null {
  const marked = all('[data-next-action] button:not([disabled]), button[data-next-action]:not([disabled])')[0] as HTMLButtonElement | undefined;
  if (marked) return marked;
  if (!row) return null;
  const edge = row.getBoundingClientRect().right - 4;
  const cands = (all('main button:not([disabled])') as HTMLButtonElement[])
    .filter((b) => DARK.test(b.className) && b.getBoundingClientRect().left >= edge);
  return cands[0] || null;
}

/** 마지막으로 고른 줄 — 고른 줄 표시(data-selected)가 없는 목록에서 방향키가 이어 갈 자리 */
let lastRow: HTMLElement | null = null;

export function ListKeys() {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.ctrlKey || e.metaKey || e.altKey || e.isComposing) return;
      if (isDialogOpen()) return;
      const t = e.target as HTMLElement | null;
      const typing = isTyping(t);

      // 검색 칸에서 ↓ = 목록 첫 줄
      if (typing) {
        if (e.key === 'ArrowDown' && t?.hasAttribute('data-list-search')) {
          const rows = all('[data-list-row]');
          if (rows.length) { e.preventDefault(); (t as HTMLInputElement).blur(); rows[0].click(); rows[0].scrollIntoView({ block: 'nearest' }); }
        }
        return;
      }
      // 버튼 · 링크에 초점이 있을 때 Enter/Space 는 그 버튼의 것
      const onControl = !!t && (t.tagName === 'BUTTON' || t.tagName === 'A' || t.getAttribute('role') === 'radio');

      if (e.key === '/') {
        const s = all('[data-list-search]')[0] as HTMLInputElement | undefined;
        if (s) { e.preventDefault(); s.focus(); s.select(); }
        return;
      }
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        if (onControl && t?.getAttribute('role') === 'radio') return; // 선택 버튼 묶음은 자기 방향키를 쓴다
        const rows = all('[data-list-row]');
        if (rows.length === 0) return;
        let cur = rows.findIndex((r) => r.getAttribute('data-selected') === 'true');
        // 고른 줄 표시가 없는 목록(상담 등)은 마지막으로 누른 줄을 기억해서 거기서 이어 간다
        if (cur < 0 && lastRow) cur = rows.indexOf(lastRow);
        const next = e.key === 'ArrowDown' ? Math.min(cur + 1, rows.length - 1) : Math.max(cur - 1, 0);
        if (next === cur) return;
        e.preventDefault();
        lastRow = rows[next];
        rows[next].click();
        rows[next].scrollIntoView({ block: 'nearest' });
        return;
      }
      if (e.key === 'Enter') {
        if (onControl) return;
        const row = all('[data-list-row][data-selected="true"]')[0] || null;
        if (!row) return;
        const btn = findNextAction(row);
        if (btn) { e.preventDefault(); btn.focus(); btn.scrollIntoView({ block: 'nearest' }); }
        return;
      }
      if (e.key === 'Escape') {
        window.dispatchEvent(new CustomEvent('mm:list-escape'));
        return;
      }
      if (e.shiftKey) return;
      if (e.key === 'n' || e.key === 'N' || e.key === 'ㅜ') {
        const b = all('[data-new-action]')[0];
        if (b) { e.preventDefault(); b.click(); }
        return;
      }
      if (/^[1-9]$/.test(e.key)) {
        const group = all('[data-status-tabs]')[0];
        const btn = group ? (Array.from(group.querySelectorAll<HTMLElement>('button')).filter(visible)[Number(e.key) - 1]) : undefined;
        if (btn) { e.preventDefault(); btn.click(); }
      }
    };
    // 마우스로 누른 줄도 기억한다
    const onClick = (e: MouseEvent) => {
      const row = (e.target as HTMLElement | null)?.closest?.('[data-list-row]');
      if (row instanceof HTMLElement) lastRow = row;
    };
    window.addEventListener('keydown', onKey);
    document.addEventListener('click', onClick, true);
    return () => { window.removeEventListener('keydown', onKey); document.removeEventListener('click', onClick, true); };
  }, []);
  return null;
}

/** 목록 화면이 Esc(상세 닫기)를 받는 곳 — 열린 상세가 없으면 아무 일도 하지 않게 화면이 판단한다 */
export function useListEscape(onEscape: () => void) {
  useEffect(() => {
    const h = () => onEscape();
    window.addEventListener('mm:list-escape', h);
    return () => window.removeEventListener('mm:list-escape', h);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  });
}
