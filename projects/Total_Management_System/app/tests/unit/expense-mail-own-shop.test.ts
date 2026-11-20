import { describe, it, expect } from 'vitest';
import { parseExpenseMail } from '@/lib/expenses/mail-parse';

// 경비 메일 판독(2026-11-20) — 우리 쇼핑몰의 주문 · 결제 알림은 매출이지 경비가 아니다
const at = '2026-11-18T03:00:00Z';

describe('경비 메일 — 우리 쇼핑몰 알림은 올리지 않는다', () => {
  it('고객 결제 완료 알림(쇼핑몰 지메일 발신)은 경비 후보가 아니다', () => {
    expect(parseExpenseMail({
      from: '"MAMORU | 미용가위 전문 브랜드" <shop@example.com>',
      subject: '[즐거운 미용가위, MAMORU] [202600000000000] 주문에 대한 결제가 완료되었습니다.',
      text: '즐거운 미용가위, MAMORU | 새로운 주문이 접수되었습니다. | 주문 정보 | 고객명 | 점검고객 | 주문번호 | 202600000000000 | 결제금액 | 280,000원',
      receivedAt: at,
    })).toBeNull();
  });
  it('제목이 달라도 본문이 주문 접수 알림이면 올리지 않는다', () => {
    expect(parseExpenseMail({
      from: 'shop@example.com', subject: '결제 확인 안내',
      text: '새로운 주문이 접수되었습니다. 결제금액 | 50,000원', receivedAt: at,
    })).toBeNull();
  });
  it('다른 곳의 영수증은 그대로 올린다', () => {
    const p = parseExpenseMail({
      from: 'Vercel Inc. <invoice@example.com>', subject: 'Your receipt from Vercel Inc. #2651-0000-0000',
      text: 'Receipt from Vercel Inc. $20.00 Paid November 17, 2026', receivedAt: at,
    });
    expect(p?.vendor).toBe('Vercel');
    expect(p?.amountOriginal).toBe(20);
  });
});
