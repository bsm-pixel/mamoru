'use client';

import { useState } from 'react';
import { Copy, Check } from 'lucide-react';

/**
 * 한 번 눌러 복사 (2026-11-15) — 전화번호 · 주소 · 송장번호 · 판매번호 옆의 작은 버튼
 * 제품 상세의 「코드 복사」 버튼과 같은 모양. 누르면 1초 동안 체크 표시로 바뀐다.
 *   <CopyButton value={s.sale_number} label="판매번호" />
 */
export function CopyButton({ value, label = '', className = '' }: { value: string | null | undefined; label?: string; className?: string }) {
  const [done, setDone] = useState(false);
  if (!value) return null;
  const copy = async (e: React.MouseEvent) => {
    e.stopPropagation();
    e.preventDefault();
    try {
      await navigator.clipboard.writeText(value);
    } catch {
      // 권한이 막힌 브라우저용 예비 방법
      const ta = document.createElement('textarea');
      ta.value = value; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); } catch { /* 무시 */ }
      document.body.removeChild(ta);
    }
    setDone(true);
    setTimeout(() => setDone(false), 1000);
  };
  const name = `${label ? label + ' ' : ''}복사`;
  return (
    <button type="button" onClick={copy} title={done ? '복사했습니다' : name} aria-label={name}
      className={`inline-flex items-center justify-center shrink-0 w-6 h-6 rounded hover:bg-neutral-200 text-neutral-400 hover:text-neutral-700 align-middle ${className}`}>
      {done ? <Check size={12} className="text-emerald-600" /> : <Copy size={12} />}
      <span role="status" className="sr-only">{done ? '복사했습니다' : ''}</span>
    </button>
  );
}
