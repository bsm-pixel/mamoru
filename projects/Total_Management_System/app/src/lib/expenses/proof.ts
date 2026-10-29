/**
 * 경비 증빙 꼬리표 (2026-10-29, 마이그 168)
 *
 * 종소세 간편장부(lib/accounting/ledger)의 「결제·증빙」 칸과 「증빙 챙길 것」 시트가 이 값을 쓴다.
 * 경비 등록 폼(app/(dashboard)/expenses)·메일 확인 대기(components/expenses/expense-inbox)·메일 수집(mail-scan)이 같이 쓴다.
 */
export type ProofType = 'card' | 'tax_invoice' | 'cash_receipt' | 'receipt' | 'none';

export const PROOF_OPTIONS: Array<{ value: ProofType; label: string; hint: string }> = [
  { value: 'card', label: '사업자카드', hint: '홈택스 사업용카드 등록분은 자동' },
  { value: 'tax_invoice', label: '세금계산서', hint: '전자 세금계산서는 홈택스 자동' },
  { value: 'cash_receipt', label: '현금영수증', hint: '지출증빙용 — 홈택스 자동' },
  { value: 'receipt', label: '영수증만', hint: '해외 결제 등 — 원본 보관' },
  { value: 'none', label: '없음', hint: '증빙 없음 — 이체 내역 등 보관' },
];

export function proofLabel(p: string | null | undefined): string {
  return PROOF_OPTIONS.find((o) => o.value === p)?.label ?? '미정';
}

/** 홈택스에 자동으로 안 잡혀 따로 챙겨야 하는 것 (미정 포함) */
export function proofNeedsKeeping(p: string | null | undefined): boolean {
  return !p || p === 'receipt' || p === 'none';
}

/** 메일 종류 → 기본 증빙 (마이그 168 의 UPDATE 와 같은 규칙) */
export function proofFromDocType(docType: string | null | undefined): ProofType {
  if (docType === 'card') return 'card';
  if (docType === 'tax_invoice') return 'tax_invoice';
  return 'receipt';
}

export function toProof(v: unknown): ProofType | null {
  return PROOF_OPTIONS.some((o) => o.value === v) ? (v as ProofType) : null;
}
