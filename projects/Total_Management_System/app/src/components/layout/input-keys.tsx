'use client';

import { useEffect } from 'react';
import { todayKST } from '@/lib/utils/kst';

/**
 * 숫자 · 날짜 칸 공통 키 (2026-11-15) — 대시보드 레이아웃에 한 번만 둔다
 *
 *   날짜 칸    T = 오늘
 *   숫자 칸    - · e 는 막는다(금액 · 수량에 음수가 들어가던 문제). 음수가 필요한 칸은 data-allow-negative
 *   수량 칸    최솟값(min)이 정해진 숫자 칸에서는 + / - 키로 1씩 올리고 내린다(최솟값 아래로는 안 내려감)
 *   큰 숫자    숫자 칸에 1,000 이상을 치면 칸 바로 위에 쉼표 붙은 값이 작게 뜬다(2026-11-16).
 *              금액 칸 50여 곳을 하나씩 바꾸지 않고 전 화면에 한 번에 적용하려고 여기에 뒀다. 끄려면 data-no-comma
 *
 * React 가 값을 알아채도록 "브라우저가 직접 넣은 것처럼" 값을 넣는다(setNative).
 */
function setNative(el: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
  setter?.call(el, value);
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
}

/** 숫자 칸 위에 뜨는 쉼표 표시 — 화면에 하나만 만들어 돌려 쓴다 */
let bubble: HTMLDivElement | null = null;
function hideBubble() {
  if (bubble) bubble.style.display = 'none';
}
function showBubble(el: HTMLInputElement) {
  const n = Number(el.value);
  if (el.type !== 'number' || el.hasAttribute('data-no-comma') || !el.value || !Number.isFinite(n) || Math.abs(n) < 1000) { hideBubble(); return; }
  if (!bubble) {
    bubble = document.createElement('div');
    bubble.setAttribute('aria-hidden', 'true');
    bubble.style.cssText = 'position:fixed;z-index:400;pointer-events:none;background:#1A1A1A;color:#fff;font-size:11px;line-height:1;padding:4px 6px;border-radius:4px;font-variant-numeric:tabular-nums;white-space:nowrap';
    document.body.appendChild(bubble);
  }
  bubble.textContent = n.toLocaleString('ko-KR');
  bubble.style.display = 'block';
  const r = el.getBoundingClientRect();
  const top = r.top - bubble.offsetHeight - 4;
  bubble.style.top = `${top < 4 ? r.bottom + 4 : top}px`;
  bubble.style.left = `${Math.max(4, r.right - bubble.offsetWidth)}px`;
}

export function InputKeys() {
  useEffect(() => {
    const onType = (e: Event) => { if (e.target instanceof HTMLInputElement && e.target.type === 'number') showBubble(e.target); };
    const onMove = () => { const el = document.activeElement; if (el instanceof HTMLInputElement && el.type === 'number') showBubble(el); else hideBubble(); };
    document.addEventListener('input', onType);
    document.addEventListener('focusin', onType);
    document.addEventListener('focusout', hideBubble);
    window.addEventListener('scroll', onMove, true);
    window.addEventListener('resize', onMove);

    const onKey = (e: KeyboardEvent) => {
      if (e.ctrlKey || e.metaKey || e.altKey || e.isComposing) return;
      const t = e.target;
      if (!(t instanceof HTMLInputElement) || t.disabled || t.readOnly) return;

      if (t.type === 'date') {
        if (e.key === 't' || e.key === 'T' || e.key === 'ㅅ') {
          const today = todayKST();
          if ((t.max && today > t.max) || (t.min && today < t.min)) return;
          e.preventDefault();
          setNative(t, today);
        }
        return;
      }

      if (t.type === 'number') {
        if (t.hasAttribute('data-allow-negative')) return;
        const minAttr = t.getAttribute('min');
        const min = minAttr !== null && minAttr !== '' ? Number(minAttr) : null;
        if (min !== null && min < 0) return; // 음수를 받는 칸
        if (e.key === 'e' || e.key === 'E') { e.preventDefault(); return; }
        if (e.key === '-' || e.key === '+' || e.key === '=') {
          e.preventDefault();
          if (min === null || (min < 1 && !t.hasAttribute('data-qty'))) return; // 금액 칸: 막기만 한다 (수량 칸 = 최솟값 1 이상 또는 data-qty)
          const step = Number(t.step) > 0 ? Number(t.step) : 1;
          const cur = Number(t.value) || 0;
          let next = e.key === '-' ? cur - step : cur + step;
          if (next < min) next = min;
          if (t.max !== '' && next > Number(t.max)) next = Number(t.max);
          if (next !== cur) setNative(t, String(next));
        }
      }
    };
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('input', onType);
      document.removeEventListener('focusin', onType);
      document.removeEventListener('focusout', hideBubble);
      window.removeEventListener('scroll', onMove, true);
      window.removeEventListener('resize', onMove);
      hideBubble();
    };
  }, []);
  return null;
}
