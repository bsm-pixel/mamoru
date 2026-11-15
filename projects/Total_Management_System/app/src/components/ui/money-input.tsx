'use client';

import type { InputHTMLAttributes } from 'react';

/**
 * 금액 칸 (2026-11-15) — 치는 동안 쉼표가 붙고, 숫자만 들어간다(음수 · 글자 불가)
 *
 * 값은 숫자(원)로 주고받는다. 빈 칸은 0.
 *   <MoneyInput value={discount} onChange={setDiscount} placeholder="0" className="…" />
 */
interface Props extends Omit<InputHTMLAttributes<HTMLInputElement>, 'value' | 'onChange' | 'type'> {
  value: number;
  onChange: (value: number) => void;
}

export function MoneyInput({ value, onChange, className = '', ...rest }: Props) {
  return (
    <input
      {...rest}
      type="text"
      inputMode="numeric"
      autoComplete="off"
      value={value ? value.toLocaleString('ko-KR') : ''}
      onChange={(e) => {
        const digits = e.target.value.replace(/[^0-9]/g, '').slice(0, 12);
        onChange(digits ? Number(digits) : 0);
      }}
      className={`${className} text-right tabular-nums`}
    />
  );
}
