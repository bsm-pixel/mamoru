import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, cancelShipment } from '@/lib/lotte/alps-client';
import { randomUUID } from 'crypto';

/**
 * GET /api/lotte/test-return-pickup?confirm=YES[&zip=&addr=&tel=&name=&code=01]
 * 롯데 ALPS "반품수거(역방향 집화)" 실측 — 보내는분=고객, 받는분=마모루로 예약이 접수되는지 확인.
 * ⚠️ 실제 롯데 접수를 시도하나, 성공 시 즉시 자동 취소한다(방문 예약 남기지 않음).
 * 로그인 필요. 결과(JSON) 전체를 복사해서 전달.
 */
export async function GET(req: NextRequest) {
  try {
    const sp = req.nextUrl.searchParams;
    // 임시 진단용: 로그인 쿠키 or 토큰 둘 중 하나 (브라우저 직접 열기용). 확인 후 이 엔드포인트 제거 예정.
    const TEST_TOKEN = 'mamoru-return-pickup-2608';
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user && sp.get('token') !== TEST_TOKEN) {
      return NextResponse.json({ error: 'Unauthorized (로그인 안 됐으면 &token=... 붙이세요)' }, { status: 401 });
    }

    const API_URL = (process.env.LOTTE_API_URL || '').trim();
    const CANCEL_URL = (process.env.LOTTE_CANCEL_API_URL || '').trim();
    const KEY = (process.env.LOTTE_CLIENT_KEY || '').trim();
    const JOB = (process.env.LOTTE_JOB_CUST_CD || process.env.LOTTE_JOBCUSTCD || '').trim();

    // 🧹 취소 모드: ?cancelInvoice=송장번호 — 방금 접수된 테스트 건을 취소 + 원시 응답 확인
    const cancelInv = (sp.get('cancelInvoice') || '').replace(/\D/g, '');
    if (cancelInv) {
      if (!CANCEL_URL) return NextResponse.json({ error: 'LOTTE_CANCEL_API_URL 미설정 — 롯데 관리자에서 수동취소 필요', invoice: cancelInv });
      const payload = { snd_list: [{ jobCustCd: JOB, invNo: cancelInv, canCd: '01', canDtlCd: '19', canRmk: '반품수거 테스트 취소(TMS)' }] };
      const res = await fetch(CANCEL_URL, {
        method: 'POST',
        headers: { Authorization: `IgtAK ${KEY}`, Accept: 'application/json', 'Content-Type': 'application/json; charset=utf-8', 'X-Idempotency-Key': randomUUID(), 'X-Correlation-Id': randomUUID() },
        body: JSON.stringify(payload),
      });
      const raw = (await res.text()).slice(0, 600);
      return NextResponse.json({ mode: 'cancel', invoice: cancelInv, httpCode: res.status, raw });
    }

    if (sp.get('confirm') !== 'YES') {
      return NextResponse.json({ error: '실측을 실행하려면 ?confirm=YES 를 붙이세요. (롯데 접수 시도 후 자동 취소)' }, { status: 400 });
    }

    const FARE = (process.env.LOTTE_DEFAULT_FARE || '03').trim();
    const SENDER = {
      name: (process.env.LOTTE_SENDER_NAME || '마모루').trim(),
      tel: (process.env.LOTTE_SENDER_TEL || '').replace(/\D/g, ''),
      zip: (process.env.LOTTE_SENDER_ZIP || '').trim(),
      addr: (process.env.LOTTE_SENDER_ADDR || '').trim(),
    };
    if (!API_URL || !KEY || !JOB) return NextResponse.json({ error: 'LOTTE 환경변수 미설정' }, { status: 500 });

    // 고객(보내는분) — 쿼리로 주면 그 주소, 없으면 마모루 자기주소로 자가 테스트
    const custZip = (sp.get('zip') || SENDER.zip).trim();
    const custAddr = (sp.get('addr') || SENDER.addr).trim();
    const custTel = (sp.get('tel') || SENDER.tel).replace(/\D/g, '');
    const custName = (sp.get('name') || '반품테스트').trim();
    const codes = (sp.get('code') || '01').split(',').map((s) => s.trim()).filter(Boolean); // ustRtgSctCd 후보(기본 01)

    const now = new Date();
    const pad2 = (n: number) => String(n).padStart(2, '0');
    const ymd = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}`;

    const results: Record<string, unknown>[] = [];
    for (const code of codes) {
      const { invoiceNumber } = await getNextInvoice();
      const ordNo = `TEST-RET-${ymd}-${pad2(now.getHours())}${pad2(now.getMinutes())}${pad2(now.getSeconds())}-${code}`;
      // 🔁 역방향: 보내는분=고객, 받는분=마모루 (롯데가 고객집 방문 수거 → 마모루 도착)
      const payload = {
        snd_list: [{
          jobCustCd: JOB, ustRtgSctCd: code, ordSct: '1', fareSctCd: FARE, ordNo, invNo: invoiceNumber,
          snperNm: custName, snperTel: custTel, snperCpno: custTel, snperZipcd: custZip, snperAdr: custAddr,
          acperNm: SENDER.name, acperTel: SENDER.tel, acperCpno: SENDER.tel, acperZipcd: SENDER.zip, acperAdr: SENDER.addr,
          boxTypCd: 'A', gdsNm: '반품수거 테스트', dlvMsgCont: '', cusMsgCont: '', pickReqYmd: ymd,
        }],
      };
      let rtnCd = '', rtnMsg = '', httpCode = 0, raw = '';
      try {
        const res = await fetch(API_URL, {
          method: 'POST',
          headers: { Authorization: `IgtAK ${KEY}`, Accept: 'application/json', 'Content-Type': 'application/json; charset=utf-8', 'X-Idempotency-Key': randomUUID(), 'X-Correlation-Id': randomUUID() },
          body: JSON.stringify(payload),
        });
        httpCode = res.status;
        raw = (await res.text()).slice(0, 600);
        let json: Record<string, unknown> = {};
        try { json = JSON.parse(raw); } catch { /* keep raw */ }
        const first = (Array.isArray(json.rtn_list) ? json.rtn_list[0] : {}) as Record<string, unknown>;
        rtnCd = String(first.rtnCd || '').toUpperCase();
        rtnMsg = String(first.rtnMsg || '');
      } catch (e) { rtnMsg = `EXCEPTION ${String(e).slice(0, 200)}`; }

      let cancelled = false;
      if (rtnCd === 'S') {
        // 성공하면 즉시 취소 — 실제 방문 예약 남기지 않음
        const c = await cancelShipment(invoiceNumber).catch(() => ({ success: false }));
        cancelled = !!c.success;
      }
      results.push({ ustRtgSctCd: code, invoiceNumber, httpCode, rtnCd, rtnMsg, booked: rtnCd === 'S', auto_cancelled: cancelled, raw });
    }

    return NextResponse.json({
      note: '보내는분=고객, 받는분=마모루(역방향)로 롯데 반품수거 접수 실측. rtnCd=S 면 접수됨(즉시 자동취소). 이 JSON 전체를 전달해주세요.',
      sender_reversed: { custZip, custAddr },
      results,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
