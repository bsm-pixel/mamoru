'use client';

import type { InputHTMLAttributes } from 'react';

/**
 * 금액 칸 (2026-11-15) — 치는 동안 쉼표가 붙고, 숫자만 들어간다(음수 · 글자 불가)
 *
 * 2026-11-16 금액 칸 전부(40여 곳)를 이 부품으로 바꿨다. 화면이 값을 어떻게 들고 있느냐에 따라 둘 중 하나를 쓴다.
 *   <MoneyInput value={discount} onChange={setDiscount} />        값이 숫자(원). 빈 칸 = 0
 *   <MoneyText  value={amountText} onChange={setAmountText} />    값이 글자("15000"). 빈 칸 = ''
 * 0 을 빈 칸이 아니라 "0" 으로 보여 줘야 하면 showZero.
 * 새 금액 칸을 만들 때 <input type="number"> 를 쓰지 않는다.
 */
type Rest = Omit<InputHTMLAttributes<HTMLInputElement>, 'value' | 'onChange' | 'type'>;

const digitsOf = (text: string) => text.replace(/[^0-9]/g, '').slice(0, 12);
const withCommas = (digits: string) => (digits ? Number(digits).toLocaleString('ko-KR') : '');

interface MoneyInputProps extends Rest {
  value: number | null | undefined;
  onChange: (value: number) => void;
  /** 0 을 빈 칸 대신 "0" 으로 보여 준다 */
  showZero?: boolean;
}

export function MoneyInput({ value, onChange, showZero = false, className = '', ...rest }: MoneyInputProps) {
  const n = typeof value === 'number' && Number.isFinite(value) ? Math.max(0, Math.trunc(value)) : 0;
  return (
    <input
      {...rest}
      type="text"
      inputMode="numeric"
      autoComplete="off"
      value={n || showZero ? n.toLocaleString('ko-KR') : ''}
      onChange={(e) => {
        const digits = digitsOf(e.target.value);
        onChange(digits ? Number(digits) : 0);
      }}
      className={`${className} text-right tabular-nums`}
    />
  );
}

interface MoneyTextProps extends Rest {
  value: string | number | null | undefined;
  /** 숫자만 남긴 글자("15000"). 다 지우면 '' */
  onChange: (value: string) => void;
}

export function MoneyText({ value, onChange, className = '', ...rest }: MoneyTextProps) {
  return (
    <input
      {...rest}
      type="text"
      inputMode="numeric"
      autoComplete="off"
      value={withCommas(digitsOf(String(value ?? '')))}
      onChange={(e) => {
        const digits = digitsOf(e.target.value);
        // 앞자리 0 은 떼어 낸다("0500" → "500"). 0 하나는 그대로
        onChange(digits ? String(Number(digits)) : '');
      }}
      className={`${className} text-right tabular-nums`}
    />
  );
}
