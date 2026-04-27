'use client';

import { useCallback } from 'react';

const SCRIPT_SRC = 'https://t1.daumcdn.net/mapjsapi/bundle/postcode/prod/postcode.v2.js';

interface DaumPostcodeData {
  zonecode: string;
  roadAddress: string;
  jibunAddress: string;
}

interface DaumWindow {
  daum?: {
    Postcode: new (opts: { oncomplete: (d: DaumPostcodeData) => void }) => { open: () => void };
  };
}

interface Props {
  onSelected: (data: { zonecode: string; roadAddress: string }) => void;
  className?: string;
  children?: React.ReactNode;
  disabled?: boolean;
}

export function DaumPostcodeButton({ onSelected, className, children, disabled }: Props) {
  const open = useCallback(() => {
    const w = window as unknown as DaumWindow;

    function doOpen() {
      const Postcode = (window as unknown as DaumWindow).daum?.Postcode;
      if (!Postcode) return;
      new Postcode({
        oncomplete: (d) => {
          onSelected({
            zonecode: d.zonecode,
            roadAddress: d.roadAddress || d.jibunAddress,
          });
        },
      }).open();
    }

    if (w.daum?.Postcode) {
      doOpen();
      return;
    }
    const existing = document.querySelector(`script[src="${SCRIPT_SRC}"]`) as HTMLScriptElement | null;
    if (existing) {
      existing.addEventListener('load', doOpen, { once: true });
      return;
    }
    const s = document.createElement('script');
    s.src = SCRIPT_SRC;
    s.async = true;
    s.onload = doOpen;
    document.head.appendChild(s);
  }, [onSelected]);

  return (
    <button
      type="button"
      onClick={open}
      disabled={disabled}
      className={
        className ??
        'h-9 px-3 rounded-lg bg-neutral-900 text-white text-xs font-medium hover:bg-neutral-800 transition disabled:opacity-50'
      }
    >
      {children ?? '주소 검색'}
    </button>
  );
}
