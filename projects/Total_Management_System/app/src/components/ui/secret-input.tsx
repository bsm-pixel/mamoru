'use client';

import { useState } from 'react';
import { Eye, EyeOff } from 'lucide-react';

/**
 * 가려 두는 입력칸 (2026-11-15) — 웹훅 주소 · 키처럼 화면에 그대로 보이면 안 되는 값
 * 기본은 가려져 있고 [보기]를 눌러야 보인다. 비어 있으면 가릴 것이 없으니 그냥 입력칸.
 */
export function SecretInput({ value, onChange, placeholder, className = '', ariaLabel }: {
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
  className?: string;
  ariaLabel?: string;
}) {
  const [show, setShow] = useState(false);
  return (
    <div className="flex items-center gap-1.5">
      <input
        type={show || !value ? 'text' : 'password'}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        aria-label={ariaLabel}
        autoComplete="off"
        spellCheck={false}
        className={`flex-1 min-w-0 h-9 px-3 rounded-lg border border-neutral-200 font-mono text-xs ${className}`}
      />
      <button type="button" onClick={() => setShow((v) => !v)} aria-pressed={show}
        className="shrink-0 h-9 px-2.5 rounded-lg border border-neutral-200 text-xs font-semibold text-neutral-600 hover:bg-neutral-50 inline-flex items-center gap-1">
        {show ? <EyeOff size={13} /> : <Eye size={13} />}{show ? '가리기' : '보기'}
      </button>
    </div>
  );
}
