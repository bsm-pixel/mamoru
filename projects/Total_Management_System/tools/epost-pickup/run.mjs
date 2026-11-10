// 인터넷우체국 방문접수(수거) 자동 접수 — 1단계 PC 시험 (2026-11-10)
// 사용: node run.mjs AS-20261110-4844 [--submit] [--show]
//   --submit 없으면 '신청' 직전까지만(검증 2개 포함). --show 는 화면 보이게(문제 볼 때).
// 순서·규칙 SSOT = .claude/skills/epost-pickup/SKILL.md (리허설 2회로 확인한 순서 그대로)
// 🔒 비밀값(로그인·카드)은 화면·로그·캡처에 남기지 않는다.
import puppeteer from 'puppeteer-core';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const AS_ID = process.argv.find((a) => /^AS-/.test(a));
const SUBMIT = process.argv.includes('--submit');
const SHOW = process.argv.includes('--show');
const CHROME = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const SECRETS = 'C:\\마모루드라이브\\ClaudeSecrets';
const TMS_ENV = path.join(HERE, '..', '..', 'app', '.env.local');
const RESERVE_URL = 'https://parcel.epost.go.kr/general.RetrieveGeneralNewGubunLogin.parcel';
const LOGIN_URL = 'https://www.epost.go.kr/usr/login/cafzc008k01.jsp';
const RECEIVER_NAME = '마모루미용가위'; // 03 받는 분 = 기존 이력 첫 행
const ITEM_NAME = '미용가위'; // 04 물품 = 불러오기

if (!AS_ID) { console.error('사용법: node run.mjs AS-YYYYMMDD-NNNN [--submit] [--show]'); process.exit(2); }

const sleep = (t) => new Promise((r) => setTimeout(r, t));
const readEnv = (file) => Object.fromEntries(
  fs.readFileSync(file, 'utf8').split(/\r?\n/).filter((l) => /^[A-Z_0-9]+=/.test(l))
    .map((l) => { const i = l.indexOf('='); return [l.slice(0, i).trim(), l.slice(i + 1).trim().replace(/^"|"$/g, '')]; }),
);

const OUT = path.join(HERE, 'out', `${AS_ID}-${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}`);
fs.mkdirSync(OUT, { recursive: true });
const logs = [];
const log = (msg) => { const line = `[${new Date().toLocaleTimeString('ko-KR')}] ${msg}`; logs.push(line); console.log(line); };
const saveLog = (extra = {}) => fs.writeFileSync(path.join(OUT, 'log.json'), JSON.stringify({ AS_ID, SUBMIT, logs, dialogs, ...extra }, null, 2));
class Stop extends Error {}
const stop = (msg) => { throw new Stop(msg); };

const dialogs = [];
let shotNo = 0;

// ── 1) TMS 복원수리 건 읽기 ──
const tms = readEnv(TMS_ENV);
const res = await fetch(`${tms.NEXT_PUBLIC_SUPABASE_URL}/rest/v1/repairs?select=as_id,name,phone,proceed_type,postcode,address,address_detail,pickup_date,delivery_method,status&as_id=eq.${encodeURIComponent(AS_ID)}`, {
  headers: { apikey: tms.SUPABASE_SERVICE_ROLE_KEY, Authorization: `Bearer ${tms.SUPABASE_SERVICE_ROLE_KEY}` },
});
const [rep] = await res.json();
if (!rep) { console.error(`TMS에 ${AS_ID} 없음`); process.exit(2); }
log(`TMS: ${rep.name} / ${rep.postcode} ${rep.address} ${rep.address_detail || ''} / 수거 ${rep.pickup_date} / ${rep.delivery_method}`);
for (const k of ['name', 'phone', 'postcode', 'address', 'pickup_date', 'delivery_method']) if (!rep[k]) { console.error(`TMS ${k} 비어 있음 → 중단`); process.exit(2); }
if (rep.proceed_type !== '방문수거') { console.error(`진행방법이 방문수거가 아님(${rep.proceed_type}) → 중단`); process.exit(2); }

const login = readEnv(path.join(SECRETS, 'epost-login.env'));
const card = readEnv(path.join(SECRETS, 'epost-card.env'));
if (!login.EPOST_ID || !login.EPOST_PW) { console.error('epost-login.env 비어 있음'); process.exit(2); }
const cn = (card.CARD_NUMBER || '').replace(/\D/g, '');
const pw2 = (card.CARD_PW2 || '').replace(/\D/g, '');
if (cn.length !== 16 || pw2.length !== 2) { console.error('epost-card.env 확인 필요'); process.exit(2); }

const phoneDigits = rep.phone.replace(/\D/g, '');
const phoneParts = phoneDigits.startsWith('02')
  ? [phoneDigits.slice(0, 2), phoneDigits.slice(2, -4), phoneDigits.slice(-4)]
  : [phoneDigits.slice(0, 3), phoneDigits.slice(3, -4), phoneDigits.slice(-4)];

// ── 브라우저 ──
const browser = await puppeteer.launch({
  executablePath: CHROME,
  headless: !SHOW,
  args: ['--lang=ko-KR', '--window-size=1400,1000', '--no-first-run', '--no-default-browser-check'],
  defaultViewport: { width: 1400, height: 1000 },
});
const ua = (await browser.userAgent()).replace('HeadlessChrome', 'Chrome');
const hooked = new WeakSet();
const hook = (p, tag) => {
  if (!p || hooked.has(p)) return; hooked.add(p);
  p.on('dialog', async (d) => {
    const msg = d.message();
    const m = `${tag}:${d.type()}:${msg}`; dialogs.push(m);
    // 보안프로그램(nProtect) 설치 권유 = [취소]가 '설치 없이 이용' → 취소 (확인 누르면 설치 페이지로 빠짐, 1차 시험 사고)
    if (d.type() === 'confirm' && /보안\s*프로그램|설치/.test(msg)) {
      log(`확인창 '취소' → ${msg.replace(/\n/g, ' ')}`);
      await d.dismiss().catch(() => {});
      return;
    }
    log(`확인창 자동확인 → ${msg.replace(/\n/g, ' ')}`);
    await d.accept().catch(() => {});
  });
};
const page = await browser.newPage();
await page.setUserAgent(ua);
page.setDefaultTimeout(20000);
hook(page, '본창');
const pops = [];
page.on('popup', (p) => { pops.push(p); hook(p, '팝업'); p.setDefaultTimeout?.(20000); });
const shot = async (name, p = page) => { shotNo += 1; await p.screenshot({ path: path.join(OUT, `${String(shotNo).padStart(2, '0')}-${name}.png`) }).catch(() => {}); };

const waitPop = async (part, since) => {
  for (let i = 0; i < 60; i++) {
    const p = pops.slice(since).find((x) => !x.isClosed() && x.url().includes(part));
    if (p) { await p.waitForFunction(() => document.readyState === 'complete').catch(() => {}); return p; }
    await sleep(250);
  }
  stop(`팝업(${part})이 안 열림`);
};
const handleByText = (p, text, { exact = true, visible = true } = {}) => p.evaluateHandle((t, ex, vis) => [...document.querySelectorAll('a,button')]
  .find((e) => (ex ? e.textContent.trim() === t : e.textContent.replace(/\s+/g, '').includes(t.replace(/\s+/g, ''))) && (!vis || e.offsetParent !== null)) || null, text, exact, visible);
const clickEl = async (h, what) => {
  if (!h || !(await h.evaluate((e) => !!e).catch(() => false))) stop(`버튼 없음: ${what}`);
  await h.evaluate((e) => e.scrollIntoView({ block: 'center' }));
  await sleep(150);
  await h.click();
};
const clickText = async (text, opt) => clickEl(await handleByText(page, text, opt), text);
// 기준 요소 뒤에 오는 첫 '다음' 버튼
const clickNextAfter = async (anchorSelectorJs) => {
  const h = await page.evaluateHandle((js) => {
    const anchor = eval(js); // eslint-disable-line no-eval
    return [...document.querySelectorAll('a')].find((a) => a.textContent.trim() === '다음' && a.offsetParent !== null
      && (anchor.compareDocumentPosition(a) & Node.DOCUMENT_POSITION_FOLLOWING)) || null;
  }, anchorSelectorJs);
  await clickEl(h, '다음');
  await sleep(1200);
};
const listRows = () => page.evaluate(() => {
  const del = [...document.querySelectorAll('a')].find((e) => e.textContent.trim() === '전체삭제');
  if (!del) return [];
  let box = del; for (let i = 0; i < 6; i++) box = box.parentElement;
  return [...box.querySelectorAll('tr')].filter((t) => /^\s*\d+\s*행/.test(t.innerText)).map((t) => t.innerText.replace(/\s+/g, ' '));
});

let resultNo = null;
try {
  // ── 2) 로그인 ──
  await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded' });
  await sleep(1500);
  if (!(await page.evaluate(() => document.body.innerText.includes('로그아웃')))) {
    await page.waitForSelector('#id');
    await page.click('#id', { clickCount: 3 }); await page.type('#id', login.EPOST_ID, { delay: 30 });
    await page.click('#passwd'); await page.type('#passwd', login.EPOST_PW, { delay: 30 });
    const btn = await page.evaluateHandle(() => {
      let box = document.querySelector('#passwd'); for (let i = 0; i < 5; i++) box = box.parentElement;
      return [...box.querySelectorAll('a,button')].find((a) => a.textContent.trim() === '로그인') || null;
    });
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded' }).catch(() => {}), clickEl(btn, '로그인')]);
    await sleep(1500);
  }
  log(`로그인 후 화면: ${page.url().replace(/\?.*$/, '')}`);
  await page.goto(RESERVE_URL, { waitUntil: 'domcontentloaded' });
  await sleep(2000);
  if (!(await page.evaluate(() => document.body.innerText.includes('로그아웃')))) { await shot('login-fail'); stop('로그인 실패(로그아웃 버튼 없음) — 차단·보안확인 가능성'); }
  if (!(await page.$('#parcelUseGuideInfoChkBox'))) { await shot('no-form'); stop('방문접수 예약 화면이 안 열림'); }
  await shot('reserve-open');
  log('방문접수소포 예약 화면 열림');

  // ── 3) 서버에 남은 받는 분 목록 비우기 ──
  const left = await listRows();
  if (left.length) {
    log(`목록에 남은 ${left.length}건 → 전체삭제`);
    await page.evaluate(() => [...document.querySelectorAll('a')].find((e) => e.textContent.trim() === '전체삭제').click());
    await sleep(2000);
    if ((await listRows()).length) stop('남은 목록 삭제 실패');
  }

  // ── 4) [필수] 체크 + 01 보내는 분 = 고객 ──
  await page.evaluate(() => { const c = document.getElementById('parcelUseGuideInfoChkBox'); if (!c.checked) c.click(); });
  await page.evaluate((v) => { const n = document.querySelector('input[name=custName]'); n.value = v; n.dispatchEvent(new Event('change', { bubbles: true })); }, rep.name);
  const zipBtn = await page.evaluateHandle(() => {
    const td = document.querySelector('input[name=newPostNum]').closest('td');
    const a = [...td.querySelectorAll('a')].find((e) => e.textContent.trim() === '주소찾기');
    if (a) a.style.display = 'inline-block'; // 숨겨져 있을 때가 있음
    return a || null;
  });
  let since = pops.length;
  await clickEl(zipBtn, '주소찾기');
  let pop = await waitPop('ZipCd', since);
  await pop.waitForSelector('#keyword');
  await pop.$eval('#keyword', (e, v) => { e.value = v; }, rep.address);
  await pop.click('#btnImgSrch');
  await pop.waitForFunction(() => document.querySelectorAll('a#hi').length > 0 || /검색결과가 없|결과가 없/.test(document.body.innerText), { timeout: 15000 }).catch(() => {});
  const picked = await pop.evaluateHandle((zip) => {
    const as = [...document.querySelectorAll('a')].filter((e) => e.offsetParent !== null);
    const pairs = [];
    as.forEach((a, i) => { if (a.id === 'hi') { let z = ''; for (let j = i - 1; j >= 0; j--) { const t = as[j].textContent.trim(); if (/^\d{5}$/.test(t)) { z = t; break; } } pairs.push({ a, z }); } });
    const m = pairs.filter((p) => p.z === zip);
    return m.length === 1 ? m[0].a : null;
  }, rep.postcode);
  if (!(await picked.evaluate((e) => !!e).catch(() => false))) { await shot('zip-nomatch', pop); stop(`주소 검색 결과 중 우편번호 ${rep.postcode} 일치가 1건이 아님`); }
  log(`주소 선택: ${await picked.evaluate((e) => e.textContent.replace(/\s+/g, ' ').trim())}`);
  await picked.click();
  await pop.waitForSelector('#reg_address2', { timeout: 10000 });
  const detail = (rep.address_detail || '').trim();
  const apt = detail.match(/(\d+)\s*동\s*(\d+)\s*호/);
  const hasApt = await pop.$('#apart_yn1');
  if (apt && hasApt) {
    await pop.click('#apart_yn1'); await sleep(200);
    await pop.$eval('#reg_apartdong', (e, v) => { e.value = v; }, apt[1]);
    await pop.$eval('#reg_apartho', (e, v) => { e.value = v; }, apt[2]);
    log(`상세주소: 아파트 ${apt[1]}동 ${apt[2]}호`);
  } else {
    if (!detail) stop('TMS 상세주소가 비어 있음 → 사람 확인 필요');
    await pop.click('#apart_yn2'); await sleep(200);
    await pop.$eval('#reg_address2', (e, v) => { e.value = v; e.dispatchEvent(new Event('change', { bubbles: true })); }, detail);
    log(`상세주소: 그대로 "${detail}"`);
  }
  const addrBtn = await handleByText(pop, '주소입력');
  await clickEl(addrBtn, '주소입력');
  await sleep(1500);
  await page.evaluate((p) => {
    const s = (id, v) => { const e = document.getElementById(id) || document.querySelector(`input[name=${id}]`); e.value = v; e.dispatchEvent(new Event('change', { bubbles: true })); };
    s('cell1', p[0]); s('cell2', p[1]); s('cell3', p[2]);
  }, phoneParts);
  const sender = await page.evaluate(() => ({ zip: document.getElementById('newPostNum').value, addr: document.getElementById('postAddr').value, detail: document.getElementById('detailAddr').value }));
  if (sender.zip !== rep.postcode) stop(`보내는 분 우편번호 불일치 ${sender.zip}`);
  log(`01 보내는 분: ${rep.name} / ${sender.zip} ${sender.addr} ${sender.detail} / ${phoneParts.join('-')}`);
  await shot('01-sender');
  await clickNextAfter("document.getElementById('cell3')");

  // ── 5) 02 방문접수 소포정보 ──
  const mr = await page.$('#mreceipt_n');
  await clickEl(mr, '모바일영수증 미신청');
  if ((await page.$eval('input[name=mreceipt]', (e) => e.value)) !== 'N') stop('모바일영수증 미신청 반영 안 됨');
  const dates = await page.$$eval('#labwishReceiptTime option', (os) => os.map((o) => o.value));
  if (!dates.includes(rep.pickup_date)) stop(`수거요청일 ${rep.pickup_date} 선택 불가 (가능: ${dates.filter(Boolean).join(', ')})`);
  await page.select('#labwishReceiptTime', rep.pickup_date);
  await sleep(800);
  await page.evaluate(() => { const t = document.getElementById('labwishReceiptTimeNm'); if (!t.value) { t.selectedIndex = 0; t.dispatchEvent(new Event('change', { bubbles: true })); } });
  await page.select('#pickupKeep', '05');
  await sleep(300);
  await page.$eval('#pickupKeepNm', (e, v) => { e.value = v; e.dispatchEvent(new Event('change', { bubbles: true })); }, rep.delivery_method);
  const pay = await page.evaluate(() => { const th = [...document.querySelectorAll('th')].find((t) => t.textContent.includes('요금부담')); const r = [...th.nextElementSibling.querySelectorAll('input[type=radio]')].find((x) => x.checked); return r ? (document.querySelector(`label[for="${r.id}"]`) || {}).textContent : ''; });
  if (!/사전결제/.test(pay || '')) stop(`요금부담이 사전결제가 아님: ${pay}`);
  log(`02 소포정보: ${pay} / 모바일영수증 미신청 / ${rep.pickup_date} / 보관장소 "${rep.delivery_method}"`);
  await shot('02-parcel');
  await clickNextAfter("document.getElementById('pickupKeepNm')");

  // ── 6) 03 받는 분 = 기존 이력 ──
  since = pops.length;
  await clickText('기존받는 분 정보검색', { exact: false });
  pop = await waitPop('RecSearchAddress', since);
  await pop.waitForFunction((n) => [...document.querySelectorAll('tr')].some((t) => t.textContent.includes(n)), {}, RECEIVER_NAME);
  const recv = await pop.evaluateHandle((n) => { const row = [...document.querySelectorAll('tr')].find((t) => t.textContent.includes(n)); return [...row.querySelectorAll('a')].find((a) => a.textContent.trim() === n) || null; }, RECEIVER_NAME);
  await clickEl(recv, RECEIVER_NAME);
  await sleep(2000);
  const recvName = await page.evaluate(() => [...document.querySelectorAll('input[type=text]')].filter((e) => e.offsetParent !== null && e.title === '이름').map((e) => e.value));
  if (!recvName.includes(RECEIVER_NAME)) stop('받는 분 적용 안 됨');
  log(`03 받는 분: ${RECEIVER_NAME} (기존 이력)`);
  await clickNextAfter("[...document.querySelectorAll('a')].find(e=>e.textContent.replace(/\\s+/g,'').includes('기존받는분정보검색'))");

  // ── 7) 04 물품정보 = 불러오기 → 목록 추가 ──
  since = pops.length;
  await clickText('물품정보불러오기', { exact: false });
  pop = await waitPop('MailContList', since);
  await pop.waitForFunction((n) => [...document.querySelectorAll('tr')].some((t) => t.textContent.includes(n)), {}, ITEM_NAME);
  const item = await pop.evaluateHandle((n) => { const row = [...document.querySelectorAll('tr')].find((t) => t.textContent.includes(n)); return [...row.querySelectorAll('a')].find((a) => a.textContent.trim() === n) || null; }, ITEM_NAME);
  await clickEl(item, ITEM_NAME);
  await sleep(2000);
  const itemInfo = await page.evaluate(() => [...document.querySelectorAll('select')].filter((s) => s.offsetParent !== null && ['중량', '크기', '내용품코드'].includes(s.title)).map((s) => `${s.title}=${(s.options[s.selectedIndex] || {}).text}`));
  log(`04 물품: ${ITEM_NAME} / ${itemInfo.join(' / ')}`);
  if ((await listRows()).length) stop('목록 추가 전에 이미 행이 있음');
  await clickText('받는 분 목록에 추가');
  await sleep(2000);
  const rows = await listRows();
  if (rows.length !== 1 || !rows[0].includes(rep.name) || !rows[0].includes(RECEIVER_NAME)) stop(`받는 분 목록이 1건이 아님/내용 불일치: ${rows.length}건`);
  log('05 목록: 정확히 1건');

  // ── 8) 05 주소검증 ──
  const before = dialogs.length;
  await clickText('주소검증');
  await sleep(3000);
  if (!dialogs.slice(before).some((m) => m.includes('모두 정상'))) { await shot('addr-verify-fail'); stop(`주소검증 실패: ${dialogs.slice(before).join(' | ')}`); }
  log('05 주소검증: 모두 정상');
  await shot('05-list');
  await clickNextAfter("[...document.querySelectorAll('a,button')].find(e=>e.textContent.trim()==='주소검증')");

  // ── 9) 06 결제수단 + 결제카드검증 (1회만) ──
  const typeInto = async (id, v) => { await page.click(`#${id}`, { clickCount: 3 }); await page.type(`#${id}`, v, { delay: 20 }); };
  await typeInto('creditNo1', cn.slice(0, 4)); await typeInto('creditNo2', cn.slice(4, 8));
  await typeInto('creditNo3', cn.slice(8, 12)); await typeInto('creditNo4', cn.slice(12, 16));
  await typeInto('creditExp1', card.CARD_EXP_MM); await typeInto('creditExp2', card.CARD_EXP_YY);
  await typeInto('creditPwd1', pw2[0]); await typeInto('creditPwd2', pw2[1]);
  await typeInto('creditBirth', card.CARD_OWNER_ID);
  const lens = await page.evaluate(() => ['creditNo1', 'creditNo2', 'creditNo3', 'creditNo4', 'creditExp1', 'creditExp2', 'creditPwd1', 'creditPwd2', 'creditBirth'].map((id) => document.getElementById(id).value.length));
  log(`06 카드 입력 자릿수: ${lens.join(',')}`);
  await clickText('결제카드검증');
  let verified = false;
  for (let i = 0; i < 20 && !verified; i++) { await sleep(500); verified = await page.evaluate(() => document.body.innerText.includes('결제카드 검증이 완료')); }
  if (!verified) stop(`결제카드검증 실패(재시도 안 함): ${dialogs.slice(-2).join(' | ')}`);
  log('06 결제카드검증: 완료');

  // ── 10) 신청 ──
  const pre = await page.evaluate(() => ({ req: document.getElementById('parcelUseGuideInfoChkBox').checked, mr: document.querySelector('input[name=mreceipt]').value }));
  if (!pre.req || pre.mr !== 'N') stop('신청 전 점검 실패');
  if (!SUBMIT) { log('🛑 --submit 없음 → 신청 직전에서 멈춤'); saveLog({ result: 'dry-ok' }); }
  else {
    log('신청 클릭');
    await Promise.all([page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => {}), clickText('신청')]);
    await sleep(3000);
    const txt = await page.evaluate(() => document.body.innerText);
    // 예약조회 표: "신청일 신청인 예약접수번호(16자리) 희망방문일 …" — 고객 이름·수거일이 같은 줄
    const line = txt.split('\n').find((l) => l.includes(rep.name) && l.includes(rep.pickup_date) && /\d{16}/.test(l)) || '';
    resultNo = (line.match(/\d{16}/) || [])[0] || null;
    log(`신청 후 화면: ${page.url().replace(/\?.*$/, '')} / 예약접수번호: ${resultNo || '못 찾음'}`);
    await shot('done');
    saveLog({ result: resultNo ? 'submitted' : 'submitted-unknown', resultNo });
  }
} catch (e) {
  const msg = e instanceof Stop ? `🛑 중단: ${e.message}` : `❌ 오류: ${e.message}`;
  log(msg);
  await shot('error');
  saveLog({ result: 'stopped', error: e.message });
  process.exitCode = 1;
} finally {
  await browser.close().catch(() => {});
  console.log(`\n결과 폴더: ${OUT}`);
}
