/**
 * 아임웹 주문 결제수단(payment_method) → TMS 결제수단 키 정규화.
 *
 * 아임웹은 결제수단을 원본 문자열로 준다(예: v2 pay_type, 웹훅 method: BANKTRANSFER/BANK_TRANSFER/CARD/vbank 등).
 * TMS 리포트/회계는 card / cash / transfer / point / phone 키로 집계하므로 맞춰준다.
 *   · 무통장입금(가상계좌 vbank)·계좌이체(bank) → transfer 로 통일
 *   · 알 수 없는 값은 원본(소문자)을 그대로 노출(집계에서 숨기지 않기 위함)
 *
 * ⚠️ 아임웹 실제 pay_type 값이 확인되면(실주문 데이터) 필요 시 매칭을 좁힐 것.
 */
export function normalizeImwebPayMethod(raw: string | null | undefined): string {
  const s = (raw || '').toString().toLowerCase().replace(/[\s_-]/g, '');
  if (!s) return 'unknown';
  if (s.includes('card') || s.includes('카드') || s.includes('신용')) return 'card';
  if (s.includes('vbank') || s.includes('virtual') || s.includes('가상') || s.includes('무통장')) return 'transfer'; // 무통장입금(가상계좌)
  if (s.includes('bank') || s.includes('transfer') || s.includes('이체') || s.includes('계좌')) return 'transfer'; // 계좌이체
  if (s.includes('cash') || s.includes('현금')) return 'cash';
  if (s.includes('point') || s.includes('적립') || s.includes('포인트')) return 'point';
  if (s.includes('phone') || s.includes('mobile') || s.includes('휴대')) return 'phone';
  return s; // 미확인 값은 원본 노출
}
