'use client';

/** 바코드 스캔 입력 — 스캐너=키보드(값+Enter). Enter 시 onScan + 자동 클리어·재포커스(연속 스캔). */

import { useEffect, useRef, useState } from 'react';
import { ScanLine } from 'lucide-react';

export function ScanInput({ onScan, placeholder = '바코드 스캔 (품목 SKU / 시리얼)', autoFocus = true }: {
  onScan: (code: string) => void;
  placeholder?: string;
  autoFocus?: boolean;
}) {
  const ref = useRef<HTMLInputElement>(null);
  const [v, setV] = useState('');

  useEffect(() => { if (autoFocus) ref.current?.focus(); }, [autoFocus]);

  return (
    <div className="relative">
      <ScanLine size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-blue-500" />
      <input
        ref={ref}
        value={v}
        onChange={(e) => setV(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            const c = v.trim();
            setV('');
            if (c) onScan(c);
            ref.current?.focus();
          }
        }}
        placeholder={placeholder}
        className="w-full h-10 pl-9 pr-3 rounded-lg border-2 border-blue-200 bg-blue-50/40 text-sm text-stone-900 placeholder:text-blue-400/70 focus:outline-none focus:ring-2 focus:ring-blue-400 transition"
      />
    </div>
  );
}
