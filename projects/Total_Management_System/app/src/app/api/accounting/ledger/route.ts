import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';
import { requireUser } from '@/lib/auth/require-user';
import { createServiceClient } from '@/lib/supabase/server';
import { buildLedger } from '@/lib/accounting/ledger';

/**
 * GET /api/accounting/ledger?year=2026 — 종소세 간편장부 엑셀 (2026-10-29)
 * 시트: 장부(거래 한 줄씩) · 월별 요약 · 증빙 챙길 것 · 안내(회계 화면과 다른 점)
 * ?format=json → 같은 내용을 JSON 으로 (대조·점검용)
 * 규칙은 lib/accounting/ledger.ts 머리말.
 */
export async function GET(req: NextRequest) {
  const deny = await requireUser(); if (deny) return deny;
  const year = parseInt(req.nextUrl.searchParams.get('year') || '', 10) || new Date().getFullYear();
  if (year < 2020 || year > 2100) return NextResponse.json({ error: 'year 범위 오류' }, { status: 400 });

  try {
    const ledger = await buildLedger(createServiceClient(), year, req.nextUrl.origin);
    if (req.nextUrl.searchParams.get('format') === 'json') return NextResponse.json(ledger);

    const wb = XLSX.utils.book_new();
    const { rows, months, totals, reconcile } = ledger;

    // 1) 장부
    const book = XLSX.utils.aoa_to_sheet([
      ['일자', '구분', '거래내용', '거래처', '수입', '비용', '결제·증빙', '비고'],
      ...rows.map((r) => [r.date, r.kind, r.content, r.party, r.income || '', r.cost || '', r.proof, r.note]),
      [],
      ['합계', '', '', '', totals.income, totals.purchases + totals.expenses, '', `남은 돈 ${(totals.income - totals.purchases - totals.expenses).toLocaleString()}원`],
    ]);
    book['!cols'] = [12, 10, 34, 18, 12, 12, 22, 40].map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(wb, book, '장부');

    // 2) 월별 요약
    const sum = XLSX.utils.aoa_to_sheet([
      ['월', '매출', '매입', '경비', '남은 돈'],
      ...months.map((m) => [m.month, m.sales, m.purchases, m.expenses, m.sales - m.purchases - m.expenses]),
      ['합계', totals.income, totals.purchases, totals.expenses, totals.income - totals.purchases - totals.expenses],
    ]);
    sum['!cols'] = [10, 14, 14, 14, 14].map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(wb, sum, '월별 요약');

    // 3) 증빙 챙길 것 — 홈택스에 자동으로 안 잡히는 비용
    const keep = rows.filter((r) => r.keep);
    const keepSheet = XLSX.utils.aoa_to_sheet([
      ['일자', '구분', '거래처', '거래내용', '금액', '챙길 것', '영수증 원본'],
      ...keep.map((r) => [r.date, r.kind, r.party, r.content, r.cost, r.keep, r.receipt ? '열기 (TMS 로그인)' : '']),
      ...(keep.length ? [] : [['', '', '', '따로 챙길 증빙이 없습니다', '', '', '']]),
    ]);
    // 보관한 영수증 = 누르면 열리는 링크 (lib/expenses/receipt-archive)
    keep.forEach((r, i) => { if (r.receipt) keepSheet[XLSX.utils.encode_cell({ r: i + 1, c: 6 })].l = { Target: r.receipt }; });
    keepSheet['!cols'] = [12, 8, 22, 30, 12, 40, 18].map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(wb, keepSheet, '증빙 챙길 것');

    // 4) 안내
    const info = XLSX.utils.aoa_to_sheet([
      [`${year}년 간편장부 — TMS 자동 생성 (${new Date(Date.now() + 9 * 3600000).toISOString().slice(0, 16).replace('T', ' ')})`],
      [],
      ['매출', '판매(판매일, 할인 뺀 금액) · 반품(반품한 날 −) · 납품 · 아임웹 주문(실결제액) · 복원수리 접수(입금일)'],
      ['비용', '매입(발주일, 원화) · 경비(경비 화면). 연간 구독은 결제한 날 전액'],
      ['날짜', '한국시간 기준'],
      [],
      ['회계 화면 매출과 다른 점'],
      ['아임웹 주문 실결제액', reconcile.onlinePaid],
      ['회계 화면 기준(상품 금액, 배송비 제외)', reconcile.onlineItems],
      ['차이 (배송비 + / 적립금·쿠폰 −)', reconcile.onlineDiff],
      ['회계 화면 방식으로 바꾼 매출 합계', reconcile.reportStyleIncome],
      ['그 밖의 차이', '회계 화면은 시각 칸을 UTC 로 잘라 1월 1일 0~9시 거래가 앞 해로 갈 수 있음 — 이 장부는 한국시간'],
      [],
      ['증빙', '사업자카드·세금계산서·현금영수증은 홈택스에 자동으로 잡힘. 「증빙 챙길 것」 시트만 따로 보관'],
    ]);
    info['!cols'] = [{ wch: 36 }, { wch: 80 }];
    XLSX.utils.book_append_sheet(wb, info, '안내');

    const buf = Buffer.from(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
    const filename = encodeURIComponent(`마모루_간편장부_${year}.xlsx`);
    return new NextResponse(new Uint8Array(buf), {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${filename}`,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
