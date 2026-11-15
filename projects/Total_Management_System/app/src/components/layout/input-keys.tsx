'use client';

import { useEffect } from 'react';
import { todayKST } from '@/lib/utils/kst';

/**
 * 숫자 · 날짜 칸 공통 키 (2026-11-15) — 대시보드 레이아웃에 한 번만 둔다
 *
 *   날짜 칸    T = 오늘
 *   숫자 칸    - · e 는 막는다(금액 · 수량에 음수가 들어가던 문제). 음수가 필요한 칸은 data-allow-negative
 *   수량 칸    최솟값(min)이 정해진 숫자 칸에서는 + / - 키로 1씩 올리고 내린다(최솟값 아래로는 안 내려감)
 *
 * React 가 값을 알아채도록 "브라우저가 직접 넣은 것처럼" 값을 넣는다(setNative).
 */
function setNative(el: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
  setter?.call(el, value);
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
}

export function InputKeys() {
  useEffect(() => {
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
    return () => document.removeEventListener('keydown', onKey);
  }, []);
  return null;
}
