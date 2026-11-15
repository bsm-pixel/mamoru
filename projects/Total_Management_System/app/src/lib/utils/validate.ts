/**
 * 입력 검사 — 화면과 서버(API)가 같은 규칙을 쓴다 (2026-11-15)
 * 각 함수는 문제가 없으면 null, 있으면 사람이 읽는 문장을 돌려준다.
 */

/** 전화번호: 비워도 됨. 국내 = 0으로 시작하는 숫자 9~12자리(02-123-4567 ~ 010-1234-5678, 안심번호 0507-1234-5678 포함), 해외 = +로 시작 8~15자리 */
export function phoneError(v: unknown): string | null {
  const raw = String(v ?? '').trim();
  if (!raw) return null;
  if (/[^\d\s\-+().]/.test(raw)) return '연락처에는 숫자와 - 만 쓸 수 있습니다';
  const digits = raw.replace(/\D/g, '');
  if (raw.startsWith('+')) return digits.length >= 8 && digits.length <= 15 ? null : '해외 번호는 +국가번호 포함 8~15자리로 입력해 주세요';
  if (!digits.startsWith('0')) return '연락처는 0으로 시작해야 합니다 (예: 010-1234-5678)';
  if (digits.length < 9 || digits.length > 12) return '연락처는 숫자 9~12자리로 입력해 주세요 (예: 010-1234-5678)';
  return null;
}

/** 이메일: 비워도 됨 */
export function emailError(v: unknown): string | null {
  const raw = String(v ?? '').trim();
  if (!raw) return null;
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(raw) ? null : '이메일 주소 모양이 올바르지 않습니다 (예: name@example.com)';
}

/** 금액·수량: 0 이상의 숫자. required 면 비워 둘 수 없다 */
export function amountError(v: unknown, label: string, opts?: { required?: boolean }): string | null {
  const empty = v === undefined || v === null || v === '';
  if (empty) return opts?.required ? `${label}을(를) 입력해 주세요` : null;
  const n = Number(v);
  if (!Number.isFinite(n)) return `${label}은(는) 숫자로 입력해 주세요`;
  if (n < 0) return `${label}은(는) 0 이상이어야 합니다`;
  return null;
}

/** 입력하는 동안 자동 하이픈 (국내 번호만. +로 시작하면 그대로 둔다) */
export function typingPhone(v: string): string {
  const raw = String(v ?? '');
  if (raw.trim().startsWith('+')) return raw;
  const d = raw.replace(/\D/g, '').slice(0, 12);
  if (d.length === 12) return `${d.slice(0, 4)}-${d.slice(4, 8)}-${d.slice(8)}`;   // 안심번호 0507-1234-5678
  if (d.startsWith('02')) {
    if (d.length <= 2) return d;
    if (d.length <= 5) return `${d.slice(0, 2)}-${d.slice(2)}`;
    if (d.length <= 9) return `${d.slice(0, 2)}-${d.slice(2, 5)}-${d.slice(5)}`;
    return `${d.slice(0, 2)}-${d.slice(2, 6)}-${d.slice(6, 10)}`;
  }
  if (d.length <= 3) return d;
  if (d.length <= 6) return `${d.slice(0, 3)}-${d.slice(3)}`;
  if (d.length <= 10) return `${d.slice(0, 3)}-${d.slice(3, 6)}-${d.slice(6)}`;
  return `${d.slice(0, 3)}-${d.slice(3, 7)}-${d.slice(7)}`;
}
