/**
 * 인터넷우체국 방문접수(수거) 자동 접수 — 서버용 (2026-11-10)
 *
 * 순서·규칙 SSOT = .claude/skills/epost-pickup/SKILL.md
 * PC 시험판 = projects/Total_Management_System/tools/epost-pickup/run.mjs (같은 순서, 2026-11-10 접수 성공 35초)
 *
 * - 보내는 분 = 고객(TMS 복원수리 건), 받는 분 = 매장(우체국 '기존받는 분' 이력), 물품 = '물품정보 불러오기'
 * - 수거일·보관장소 문구는 그 건의 pickup_date / delivery_method 그대로 (대체값 금지 — 안 맞으면 중단)
 * - 확인창: 보안프로그램 설치 권유만 [취소], 나머지 [확인]
 * - 신청 전 필수 3가지: [필수] 체크 · 주소검증 '모두 정상' · 결제카드검증 완료
 * 🔒 로그인·카드 값은 로그에 남기지 않는다 (자릿수만)
 */
import type { Browser, ElementHandle, JSHandle, Page } from 'puppeteer-core';

export interface EpostRepair {
  as_id: string;
  name: string;
  phone: string;
  postcode: string;
  address: string;
  address_detail: string | null;
  pickup_date: string;
  delivery_method: string;
}
export interface EpostSecrets {
  id: string;
  pw: string;
  cardNumber: string; // 16자리
  cardExpMM: string;
  cardExpYY: string;
  cardPw2: string; // 비밀번호 앞 2자리
  cardOwnerId: string; // 생년월일 6자리 또는 사업자번호 10자리
}
export interface EpostResult {
  ok: boolean;
  submitted: boolean;
  reservationNo: string | null;
  stoppedReason: string | null;
  logs: string[];
  dialogs: string[];
  errorShot: string | null; // 실패 시 화면 (jpeg base64) — 06 결제 단계 이후엔 남기지 않음
  ms: number;
}

const LOGIN_URL = 'https://www.epost.go.kr/usr/login/cafzc008k01.jsp';
const RESERVE_URL = 'https://parcel.epost.go.kr/general.RetrieveGeneralNewGubunLogin.parcel';
const RECEIVER_NAME = '마모루미용가위';
const ITEM_NAME = '미용가위';
const IS_LOGGED_IN_JS = "[...document.querySelectorAll('a,button')].some((e) => e.textContent.trim() === '로그아웃')";
const POLL = { polling: 250 } as const;

class Stop extends Error {}
const sleep = (t: number) => new Promise((r) => setTimeout(r, t));

/** 입력값 사전 점검 — 하나라도 비면 우체국에 들어가지 않는다 */
export function validateRepairForEpost(r: Partial<EpostRepair> & { proceed_type?: string | null }): string | null {
  if (r.proceed_type && r.proceed_type !== '방문수거') return `진행방법이 방문수거가 아님(${r.proceed_type})`;
  for (const k of ['name', 'phone', 'postcode', 'address', 'address_detail', 'pickup_date', 'delivery_method'] as const) {
    if (!r[k] || !String(r[k]).trim()) return `복원수리 건의 ${k} 가 비어 있음`;
  }
  return null;
}

export async function runEpostPickup(opts: {
  browser: Browser;
  repair: EpostRepair;
  secrets: EpostSecrets;
  submit: boolean;
}): Promise<EpostResult> {
  const { browser, repair: rep, secrets: sec, submit } = opts;
  const started = Date.now();
  const logs: string[] = [];
  const dialogs: string[] = [];
  let errorShot: string | null = null;
  let cardStage = false;
  const log = (m: string) => { logs.push(`${((Date.now() - started) / 1000).toFixed(1)}s ${m}`); };
  const stop = (m: string): never => { throw new Stop(m); };
  const T = <V>(label: string, p: Promise<V>, ms = 20000): Promise<V> => {
    log(`· ${label}`);
    return Promise.race([p, sleep(ms).then(() => stop(`${label} — ${ms / 1000}초 넘게 응답 없음`))]);
  };

  const cn = sec.cardNumber.replace(/\D/g, '');
  const pw2 = sec.cardPw2.replace(/\D/g, '');
  const digits = rep.phone.replace(/\D/g, '');
  const phoneParts = digits.startsWith('02')
    ? [digits.slice(0, 2), digits.slice(2, -4), digits.slice(-4)]
    : [digits.slice(0, 3), digits.slice(3, -4), digits.slice(-4)];

  const page = await browser.newPage();
  const ua = (await browser.userAgent()).replace('HeadlessChrome', 'Chrome');
  await page.setUserAgent(ua);
  page.setDefaultTimeout(20000);
  await page.setViewport({ width: 1400, height: 1000 });

  const hooked = new WeakSet<Page>();
  const hook = (p: Page, tag: string) => {
    if (hooked.has(p)) return;
    hooked.add(p);
    p.on('dialog', async (d) => {
      const msg = d.message();
      dialogs.push(`${tag}:${d.type()}:${msg}`);
      if (d.type() === 'confirm' && /보안\s*프로그램|설치/.test(msg)) { log(`확인창 취소: ${msg}`); await d.dismiss().catch(() => {}); return; }
      log(`확인창 확인: ${msg.replace(/\n/g, ' ')}`);
      await d.accept().catch(() => {});
    });
  };
  hook(page, '본창');
  const pops: Page[] = [];
  page.on('popup', (p) => { if (p) { pops.push(p); hook(p, '팝업'); p.setDefaultTimeout(20000); } });

  const waitPop = async (part: string, since: number): Promise<Page> => {
    for (let i = 0; i < 60; i++) {
      const p = pops.slice(since).find((x) => !x.isClosed() && x.url().includes(part));
      if (p) { await p.waitForFunction(() => document.readyState === 'complete', { polling: 250, timeout: 15000 }).catch(() => {}); return p; }
      await sleep(250);
    }
    return stop(`팝업(${part})이 안 열림`);
  };
  const asEl = (h: JSHandle): ElementHandle<Element> | null => h.asElement() as ElementHandle<Element> | null;
  const byText = async (p: Page, text: string, exact = true) => asEl(await p.evaluateHandle((t: string, ex: boolean) => [...document.querySelectorAll('a,button')]
    .find((e) => (ex ? (e.textContent || '').trim() === t : (e.textContent || '').replace(/\s+/g, '').includes(t.replace(/\s+/g, ''))) && (e as HTMLElement).offsetParent !== null) || null, text, exact));
  // 실제 마우스 클릭 — 팝업을 여는 버튼·모바일영수증 체크박스용
  const realClick = async (h: ElementHandle<Element> | null, what: string) => {
    if (!h) stop(`버튼 없음: ${what}`);
    await h!.evaluate((e) => e.scrollIntoView({ block: 'center' }));
    await sleep(150);
    await h!.click();
  };
  // 화면 안 직접 클릭(비동기) — 팝업 안 링크·숨은 라디오·'다음'용
  const domClick = async (h: ElementHandle<Element> | null, what: string) => {
    if (!h) stop(`버튼 없음: ${what}`);
    await h!.evaluate((e) => { setTimeout(() => (e as HTMLElement).click(), 0); });
  };
  const clickNextAfter = async (anchorJs: string) => {
    const h = asEl(await page.evaluateHandle((js: string) => {
      const anchor = (0, eval)(js) as Element; // 기준 요소 뒤 첫 '다음'
      return [...document.querySelectorAll('a')].find((a) => a.textContent?.trim() === '다음' && (a as HTMLElement).offsetParent !== null
        && (anchor.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING)) || null;
    }, anchorJs));
    await realClick(h, '다음');
    await sleep(1200);
  };
  const listRows = () => page.evaluate(() => {
    const del = [...document.querySelectorAll('a')].find((e) => e.textContent?.trim() === '전체삭제');
    if (!del) return [] as string[];
    let box: Element = del; for (let i = 0; i < 6 && box.parentElement; i++) box = box.parentElement;
    return [...box.querySelectorAll('tr')].filter((t) => /^\s*\d+\s*행/.test((t as HTMLElement).innerText)).map((t) => (t as HTMLElement).innerText.replace(/\s+/g, ' '));
  });

  let submitted = false;
  let reservationNo: string | null = null;
  try {
    // ── 로그인 ──
    await T('로그인 화면', page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' }));
    await sleep(1500);
    if (!(await page.evaluate(IS_LOGGED_IN_JS))) {
      await page.waitForSelector('#id');
      await page.$eval('#id', (e) => { (e as HTMLInputElement).value = ''; });
      await page.click('#id'); await page.type('#id', sec.id, { delay: 20 });
      await page.click('#passwd'); await page.type('#passwd', sec.pw, { delay: 20 });
      const btn = asEl(await page.evaluateHandle(() => {
        let box: Element = document.querySelector('#passwd')!; for (let i = 0; i < 5 && box.parentElement; i++) box = box.parentElement;
        return [...box.querySelectorAll('a,button')].find((a) => a.textContent?.trim() === '로그인') || null;
      }));
      await realClick(btn, '로그인'); // checkVal() → NetFunnel(접속 대기열) → submit
      let ok = false;
      for (let i = 0; i < 30 && !ok; i++) {
        await sleep(500);
        ok = await page.evaluate(`!location.pathname.includes('/usr/login/') || ${IS_LOGGED_IN_JS}`).catch(() => false) as boolean;
      }
      if (!ok) stop('로그인 버튼 후 화면이 안 넘어감 (차단·대기열·비밀번호 확인 필요)');
      await sleep(1500);
    }
    await T('방문접수 예약 화면', page.goto(RESERVE_URL, { waitUntil: 'domcontentloaded' }));
    await sleep(2000);
    if (!(await page.evaluate(IS_LOGGED_IN_JS))) stop('로그인 실패(로그아웃 버튼 없음)');
    if (!(await page.$('#parcelUseGuideInfoChkBox'))) stop('방문접수 예약 화면이 안 열림');
    log('예약 화면 열림');

    // ── 서버에 남은 받는 분 목록 비우기 ──
    if ((await listRows()).length) {
      log('남은 목록 → 전체삭제');
      await page.evaluate(() => { ([...document.querySelectorAll('a')].find((e) => e.textContent?.trim() === '전체삭제') as HTMLElement).click(); });
      await sleep(2000);
      if ((await listRows()).length) stop('남은 목록 삭제 실패');
    }

    // ── 01 보내는 분 = 고객 ──
    await page.evaluate((name: string) => {
      const c = document.getElementById('parcelUseGuideInfoChkBox') as HTMLInputElement; if (!c.checked) c.click();
      const n = document.querySelector('input[name=custName]') as HTMLInputElement; n.value = name; n.dispatchEvent(new Event('change', { bubbles: true }));
    }, rep.name);
    const zipBtn = asEl(await page.evaluateHandle(() => {
      const td = document.querySelector('input[name=newPostNum]')!.closest('td')!;
      const a = [...td.querySelectorAll('a')].find((e) => e.textContent?.trim() === '주소찾기') as HTMLElement | undefined;
      if (a) a.style.display = 'inline-block'; // 숨겨져 있을 때가 있음
      return a || null;
    }));
    let since = pops.length;
    await T('주소찾기', realClick(zipBtn, '주소찾기'));
    let pop = await T('주소찾기 팝업', waitPop('ZipCd', since));
    await T('검색칸', pop.waitForSelector('#keyword'));
    await pop.$eval('#keyword', (e, v) => { (e as HTMLInputElement).value = v as string; }, rep.address);
    await T('검색', domClick(await pop.$('#btnImgSrch'), '검색'));
    await T('검색 결과', pop.waitForFunction(() => document.querySelectorAll('a#hi').length > 0 || /결과가 없/.test(document.body.innerText), { ...POLL, timeout: 15000 }).then(() => true).catch(() => false));
    const picked = asEl(await pop.evaluateHandle((zip: string) => {
      const as = [...document.querySelectorAll('a')].filter((e) => (e as HTMLElement).offsetParent !== null);
      const pairs: { a: Element; z: string }[] = [];
      as.forEach((a, i) => { if (a.id === 'hi') { let z = ''; for (let j = i - 1; j >= 0; j--) { const t = (as[j].textContent || '').trim(); if (/^\d{5}$/.test(t)) { z = t; break; } } pairs.push({ a, z }); } });
      const m = pairs.filter((p) => p.z === zip);
      return m.length === 1 ? m[0].a : null;
    }, rep.postcode));
    if (!picked) stop(`주소 검색 결과 중 우편번호 ${rep.postcode} 일치가 1건이 아님`);
    log(`주소 선택: ${await picked!.evaluate((e) => (e.textContent || '').replace(/\s+/g, ' ').trim())}`);
    await T('주소 클릭', domClick(picked, '주소'));
    await T('상세주소 칸', pop.waitForFunction(() => { const e = document.getElementById('reg_address2') as HTMLElement | null; return !!e && e.offsetParent !== null; }, { ...POLL, timeout: 10000 }));
    const detail = (rep.address_detail || '').trim();
    const apt = detail.match(/(\d+)\s*동\s*(\d+)\s*호/);
    if (apt && (await pop.$('#apart_yn1'))) {
      await pop.evaluate((d: string, h: string) => {
        (document.getElementById('apart_yn1') as HTMLElement).click();
        (document.getElementById('reg_apartdong') as HTMLInputElement).value = d;
        (document.getElementById('reg_apartho') as HTMLInputElement).value = h;
      }, apt[1], apt[2]);
      log(`상세주소: 아파트 ${apt[1]}동 ${apt[2]}호`);
    } else {
      await pop.evaluate((v: string) => {
        (document.getElementById('apart_yn2') as HTMLElement).click();
        const f = document.getElementById('reg_address2') as HTMLInputElement; f.value = v; f.dispatchEvent(new Event('change', { bubbles: true }));
      }, detail);
      log(`상세주소: "${detail}"`);
    }
    await sleep(300);
    await T('주소입력', domClick(await byText(pop, '주소입력'), '주소입력'));
    for (let i = 0; i < 40 && !pop.isClosed(); i++) await sleep(250);
    await sleep(800);
    await page.evaluate((p: string[]) => {
      const s = (id: string, v: string) => { const e = (document.getElementById(id) || document.querySelector(`input[name=${id}]`)) as HTMLInputElement; e.value = v; e.dispatchEvent(new Event('change', { bubbles: true })); };
      s('cell1', p[0]); s('cell2', p[1]); s('cell3', p[2]);
    }, phoneParts);
    const sender = await page.evaluate(() => ({
      zip: (document.getElementById('newPostNum') as HTMLInputElement).value,
      addr: (document.getElementById('postAddr') as HTMLInputElement).value,
      detail: (document.getElementById('detailAddr') as HTMLInputElement).value,
    }));
    if (sender.zip !== rep.postcode) stop(`보내는 분 우편번호 불일치 ${sender.zip}`);
    log(`01 보내는 분: ${rep.name} / ${sender.zip} ${sender.addr} ${sender.detail}`);
    await clickNextAfter("document.getElementById('cell3')");

    // ── 02 방문접수 소포정보 ──
    await T('모바일영수증 미신청', realClick(await page.$('#mreceipt_n'), '모바일영수증 미신청'));
    if ((await page.$eval('input[name=mreceipt]', (e) => (e as HTMLInputElement).value)) !== 'N') stop('모바일영수증 미신청 반영 안 됨');
    const dates = await page.$$eval('#labwishReceiptTime option', (os) => os.map((o) => (o as HTMLOptionElement).value));
    if (!dates.includes(rep.pickup_date)) stop(`수거요청일 ${rep.pickup_date} 선택 불가 (가능: ${dates.filter(Boolean).join(', ')})`);
    await page.select('#labwishReceiptTime', rep.pickup_date);
    await sleep(800);
    await page.evaluate(() => { const t = document.getElementById('labwishReceiptTimeNm') as HTMLSelectElement; if (!t.value) { t.selectedIndex = 0; t.dispatchEvent(new Event('change', { bubbles: true })); } });
    await page.select('#pickupKeep', '05'); // 직접입력
    await sleep(300);
    await page.$eval('#pickupKeepNm', (e, v) => { (e as HTMLInputElement).value = v as string; e.dispatchEvent(new Event('change', { bubbles: true })); }, rep.delivery_method);
    const pay = await page.evaluate(() => {
      const th = [...document.querySelectorAll('th')].find((t) => t.textContent?.includes('요금부담'));
      const r = th ? [...th.nextElementSibling!.querySelectorAll('input[type=radio]')].find((x) => (x as HTMLInputElement).checked) : null;
      return r ? (document.querySelector(`label[for="${r.id}"]`)?.textContent || '') : '';
    });
    if (!/사전결제/.test(pay)) stop(`요금부담이 사전결제가 아님: ${pay}`);
    log(`02 소포정보: ${pay} / 모바일영수증 미신청 / ${rep.pickup_date} / "${rep.delivery_method}"`);
    await clickNextAfter("document.getElementById('pickupKeepNm')");

    // ── 03 받는 분 = 기존 이력 ──
    since = pops.length;
    await T('기존받는 분 정보검색', realClick(await byText(page, '기존받는 분 정보검색', false), '기존받는 분 정보검색'));
    pop = await T('받는 분 팝업', waitPop('RecSearchAddress', since));
    await T('받는 분 이력', pop.waitForFunction((n: string) => [...document.querySelectorAll('tr')].some((t) => t.textContent?.includes(n)), POLL, RECEIVER_NAME));
    const recv = asEl(await pop.evaluateHandle((n: string) => { const row = [...document.querySelectorAll('tr')].find((t) => t.textContent?.includes(n))!; return [...row.querySelectorAll('a')].find((a) => a.textContent?.trim() === n) || null; }, RECEIVER_NAME));
    await T('받는 분 선택', domClick(recv, RECEIVER_NAME));
    await sleep(2000);
    const recvNames = await page.evaluate(() => [...document.querySelectorAll('input[type=text]')].filter((e) => (e as HTMLElement).offsetParent !== null && (e as HTMLInputElement).title === '이름').map((e) => (e as HTMLInputElement).value));
    if (!recvNames.includes(RECEIVER_NAME)) stop('받는 분 적용 안 됨');
    log(`03 받는 분: ${RECEIVER_NAME}`);
    await clickNextAfter("[...document.querySelectorAll('a')].find(e=>e.textContent.replace(/\\s+/g,'').includes('기존받는분정보검색'))");

    // ── 04 물품 = 불러오기 → 목록 추가 ──
    since = pops.length;
    await T('물품정보 불러오기', realClick(await byText(page, '물품정보불러오기', false), '물품정보 불러오기'));
    pop = await T('물품 팝업', waitPop('MailContList', since));
    await T('물품 목록', pop.waitForFunction((n: string) => [...document.querySelectorAll('tr')].some((t) => t.textContent?.includes(n)), POLL, ITEM_NAME));
    const item = asEl(await pop.evaluateHandle((n: string) => { const row = [...document.querySelectorAll('tr')].find((t) => t.textContent?.includes(n))!; return [...row.querySelectorAll('a')].find((a) => a.textContent?.trim() === n) || null; }, ITEM_NAME));
    await T('물품 선택', domClick(item, ITEM_NAME));
    await sleep(2000);
    if ((await listRows()).length) stop('목록 추가 전에 이미 행이 있음');
    await T('받는 분 목록에 추가', realClick(await byText(page, '받는 분 목록에 추가'), '받는 분 목록에 추가'));
    await sleep(2000);
    const rows = await listRows();
    if (rows.length !== 1 || !rows[0].includes(rep.name) || !rows[0].includes(RECEIVER_NAME)) stop(`받는 분 목록이 정확히 1건이 아님(${rows.length}건)`);
    log('05 목록: 정확히 1건');

    // ── 05 주소검증 ──
    const before = dialogs.length;
    await T('주소검증', realClick(await byText(page, '주소검증'), '주소검증'));
    await sleep(3000);
    if (!dialogs.slice(before).some((m) => m.includes('모두 정상'))) stop(`주소검증 실패: ${dialogs.slice(before).join(' | ')}`);
    log('05 주소검증: 모두 정상');
    await page.evaluate(() => { const nx = [...document.querySelectorAll('a')].filter((a) => a.textContent?.trim() === '다음' && (a as HTMLElement).offsetParent !== null); nx[nx.length - 1].click(); });
    const cardVisible = () => page.evaluate(() => { const e = document.getElementById('creditNo1') as HTMLElement | null; return !!e && e.offsetParent !== null && e.getBoundingClientRect().height > 0; });
    for (let i = 0; i < 12 && !(await cardVisible()); i++) await sleep(250);
    if (!(await cardVisible())) {
      await page.evaluate(() => { const h = [...document.querySelectorAll('*')].find((e) => e.childElementCount < 4 && /^06\s*결제수단\s*등록/.test((e.textContent || '').trim())) as HTMLElement | undefined; ((h?.closest('a,button,h3,h4,div') as HTMLElement) || h)?.click(); });
      for (let i = 0; i < 12 && !(await cardVisible()); i++) await sleep(250);
    }
    if (!(await cardVisible())) stop('06 결제수단 칸이 안 열림');

    // ── 06 결제수단 + 결제카드검증 (1회만, 재시도 금지 — 3회 틀리면 카드 정지) ──
    cardStage = true;
    const typeInto = async (id: string, v: string) => {
      await page.evaluate((x: string) => { const e = document.getElementById(x) as HTMLInputElement; e.scrollIntoView({ block: 'center' }); e.value = ''; e.focus(); }, id);
      await page.keyboard.type(v, { delay: 15 });
    };
    await typeInto('creditNo1', cn.slice(0, 4)); await typeInto('creditNo2', cn.slice(4, 8));
    await typeInto('creditNo3', cn.slice(8, 12)); await typeInto('creditNo4', cn.slice(12, 16));
    await typeInto('creditExp1', sec.cardExpMM); await typeInto('creditExp2', sec.cardExpYY);
    await typeInto('creditPwd1', pw2[0]); await typeInto('creditPwd2', pw2[1]);
    await typeInto('creditBirth', sec.cardOwnerId);
    const lens = await page.evaluate(() => ['creditNo1', 'creditNo2', 'creditNo3', 'creditNo4', 'creditExp1', 'creditExp2', 'creditPwd1', 'creditPwd2', 'creditBirth'].map((id) => (document.getElementById(id) as HTMLInputElement).value.length));
    log(`06 카드 자릿수: ${lens.join(',')}`);
    await T('결제카드검증', domClick(await byText(page, '결제카드검증'), '결제카드검증'));
    let verified = false;
    for (let i = 0; i < 20 && !verified; i++) { await sleep(500); verified = await page.evaluate(() => document.body.innerText.includes('결제카드 검증이 완료')); }
    if (!verified) stop(`결제카드검증 실패(재시도 안 함): ${dialogs.slice(-2).join(' | ')}`);
    log('06 결제카드검증: 완료');

    const pre = await page.evaluate(() => ({
      req: (document.getElementById('parcelUseGuideInfoChkBox') as HTMLInputElement).checked,
      mr: (document.querySelector('input[name=mreceipt]') as HTMLInputElement).value,
    }));
    if (!pre.req || pre.mr !== 'N') stop('신청 전 점검 실패');

    // ── 신청 ──
    if (!submit) {
      log('신청 직전에서 멈춤(시험 모드)');
    } else {
      log('신청');
      await Promise.all([
        page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => {}),
        realClick(await byText(page, '신청'), '신청'),
      ]);
      await sleep(3000);
      submitted = dialogs.some((m) => m.includes('성공적으로 예약'));
      const txt = await page.evaluate(() => document.body.innerText);
      const line = txt.split('\n').find((l) => l.includes(rep.name) && l.includes(rep.pickup_date) && /\d{16}/.test(l)) || '';
      reservationNo = (line.match(/\d{16}/) || [])[0] || null;
      log(`신청 결과: ${submitted ? '예약됨' : '예약 확인창 없음'} / 예약접수번호 ${reservationNo || '못 찾음'}`);
      if (!submitted && !reservationNo) stop('신청 후 예약 완료를 확인하지 못함 — 우체국 예약조회에서 직접 확인 필요');
    }
    return { ok: true, submitted, reservationNo, stoppedReason: null, logs, dialogs, errorShot: null, ms: Date.now() - started };
  } catch (e) {
    const msg = e instanceof Stop ? e.message : `오류: ${e instanceof Error ? e.message : String(e)}`;
    log(`중단: ${msg}`);
    if (!cardStage) { // 카드 입력 단계 이후 화면은 남기지 않는다
      errorShot = await page.screenshot({ type: 'jpeg', quality: 50, encoding: 'base64' }).catch(() => null) as string | null;
    }
    return { ok: false, submitted, reservationNo, stoppedReason: msg, logs, dialogs, errorShot, ms: Date.now() - started };
  } finally {
    await page.close().catch(() => {});
  }
}

/** 서버(Vercel)·로컬 공통 브라우저 실행 */
export async function launchEpostBrowser(): Promise<Browser> {
  const puppeteer = (await import('puppeteer-core')).default;
  if (process.env.VERCEL) {
    const chromium = (await import('@sparticuz/chromium')).default;
    chromium.setGraphicsMode = false; // WebGL 불필요
    return puppeteer.launch({
      args: await puppeteer.defaultArgs({ args: [...chromium.args, '--lang=ko-KR'], headless: 'shell' }),
      executablePath: await chromium.executablePath(),
      headless: 'shell',
    });
  }
  return puppeteer.launch({
    executablePath: process.env.CHROME_PATH || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    headless: true,
    args: ['--lang=ko-KR'],
  });
}

/** 환경변수 → 비밀값 (없으면 null) */
export function epostSecretsFromEnv(): EpostSecrets | null {
  const e = process.env;
  const s: EpostSecrets = {
    id: e.EPOST_ID || '', pw: e.EPOST_PW || '',
    cardNumber: e.EPOST_CARD_NUMBER || '', cardExpMM: e.EPOST_CARD_EXP_MM || '', cardExpYY: e.EPOST_CARD_EXP_YY || '',
    cardPw2: e.EPOST_CARD_PW2 || '', cardOwnerId: e.EPOST_CARD_OWNER_ID || '',
  };
  if (!s.id || !s.pw || s.cardNumber.replace(/\D/g, '').length !== 16 || s.cardPw2.replace(/\D/g, '').length !== 2 || !s.cardExpMM || !s.cardExpYY || !s.cardOwnerId) return null;
  return s;
}
