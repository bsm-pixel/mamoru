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
const MANUAL_LOGIN = process.argv.includes('--manual-login'); // 사람이 직접 로그인(창 표시 강제)
const SHOW = process.argv.includes('--show') || MANUAL_LOGIN;
const CHROME = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const SECRETS = 'C:\\마모루드라이브\\ClaudeSecrets';
const TMS_ENV = path.join(HERE, '..', '..', 'app', '.env.local');
const RESERVE_URL = 'https://parcel.epost.go.kr/general.RetrieveGeneralNewGubunLogin.parcel';
const LOGIN_URL = 'https://www.epost.go.kr/usr/login/cafzc008k01.jsp';
const RECEIVER_NAME = '마모루미용가위'; // 03 받는 분 = 기존 이력 첫 행
const ITEM_NAME = '미용가위'; // 04 물품 = 불러오기

if (!AS_ID) { console.error('사용법: node run.mjs AS-YYYYMMDD-NNNN [--submit] [--show]'); process.exit(2); }

// 로그인 여부 = '로그아웃' 링크(버튼)가 있는지. 본문 글자로 보면 안 됨 — 로그인 화면 안내문에 '자동 로그아웃' 문구가 있음(3차 시험 오판)
const IS_LOGGED_IN_JS = "[...document.querySelectorAll('a,button')].some((e) => e.textContent.trim() === '로그아웃')";
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

const login = MANUAL_LOGIN ? {} : readEnv(path.join(SECRETS, 'epost-login.env')); // 직접 로그인 모드에선 비밀번호 파일 안 읽음
const card = readEnv(path.join(SECRETS, 'epost-card.env'));
if (!MANUAL_LOGIN && (!login.EPOST_ID || !login.EPOST_PW)) { console.error('epost-login.env 비어 있음'); process.exit(2); }
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
    if (p) { await p.waitForFunction(() => document.readyState === 'complete', { polling: 250, timeout: 15000 }).catch(() => {}); return p; }
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
// 단계 시간 제한 — 어디서 멈췄는지 바로 찍히게 (4차 시험: 주소찾기 팝업에서 무한 대기)
const T = (label, p, ms = 20000) => { log(`  · ${label}`); return Promise.race([p, sleep(ms).then(() => stop(`${label} — ${ms / 1000}초 넘게 응답 없음`))]); };
// 화면 안 직접 클릭(비동기) — 팝업 안 링크·숨은 라디오용. 확인창은 puppeteer 가 처리
const domClick = (h) => h.evaluate((e) => { setTimeout(() => e.click(), 0); });
const POLL = { polling: 250 };
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
  if (MANUAL_LOGIN && !(await page.evaluate(IS_LOGGED_IN_JS))) {
    // 사람이 뜬 창에서 직접 로그인 → 로그인 확인되면 이어서 자동 진행 (Claude 실행 시 비밀번호 입력 금지)
    log('👉 크롬 창에서 직접 로그인해 주세요 (최대 5분 대기)');
    let ok = false;
    for (let i = 0; i < 600 && !ok; i++) { await sleep(500); ok = await page.evaluate(IS_LOGGED_IN_JS).catch(() => false); }
    if (!ok) stop('5분 안에 로그인 확인 안 됨');
    log('로그인 확인됨 → 자동 진행');
  } else if (!(await page.evaluate(IS_LOGGED_IN_JS))) {
    await page.waitForSelector('#id');
    await page.click('#id', { clickCount: 3 }); await page.type('#id', login.EPOST_ID, { delay: 30 });
    await page.click('#passwd'); await page.type('#passwd', login.EPOST_PW, { delay: 30 });
    const btn = await page.evaluateHandle(() => {
      let box = document.querySelector('#passwd'); for (let i = 0; i < 5; i++) box = box.parentElement;
      return [...box.querySelectorAll('a,button')].find((a) => a.textContent.trim() === '로그인') || null;
    });
    const idLen = await page.$eval('#id', (e) => e.value.length);
    const pwLen = await page.$eval('#passwd', (e) => e.value.length);
    log(`로그인 입력 자릿수: 아이디 ${idLen} / 비밀번호 ${pwLen}`);
    await shot('login-filled-masked'); // 비밀번호 칸은 화면에서 ●로 가려짐
    await clickEl(btn, '로그인'); // checkVal() → NetFunnel(접속 대기열) → submit
    // 로그인은 대기열을 거쳐 비동기로 넘어감 → 화면이 바뀌거나 '로그아웃'이 보일 때까지 최대 15초
    let ok = false;
    for (let i = 0; i < 30 && !ok; i++) {
      await sleep(500);
      ok = await page.evaluate("!location.pathname.includes('/usr/login/') || " + IS_LOGGED_IN_JS).catch(() => false);
    }
    if (!ok) {
      // 화면 안 안내창(cfAlertMsg 레이어·대기열 창) 문구 수집
      const layer = await page.evaluate(() => [...document.querySelectorAll('div,section,p')]
        .filter((e) => e.offsetParent !== null && e.childElementCount < 6 && /잠시|중단|대기|접속자|확인|오류|일치|잘못|비밀번호/.test(e.innerText) && e.innerText.length < 200)
        .map((e) => e.innerText.replace(/\s+/g, ' ').trim()).filter((t, i, a) => a.indexOf(t) === i).slice(0, 8)).catch(() => []);
      await shot('login-stuck');
      stop(`로그인 버튼 후 화면이 안 넘어감. 화면 문구: ${layer.join(' | ') || '(없음)'}`);
    }
    await sleep(1500);
  }
  log(`로그인 후 화면: ${page.url().replace(/\?.*$/, '')}`);
  await page.goto(RESERVE_URL, { waitUntil: 'domcontentloaded' });
  await sleep(2000);
  if (!(await page.evaluate(IS_LOGGED_IN_JS))) { await shot('login-fail'); stop('로그인 실패(로그아웃 버튼 없음) — 차단·보안확인 가능성'); }
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
  await T('주소찾기 클릭', clickEl(zipBtn, '주소찾기'));
  let pop = await T('주소찾기 팝업 대기', waitPop('ZipCd', since));
  await T('검색칸 대기', pop.waitForSelector('#keyword'));
  await pop.$eval('#keyword', (e, v) => { e.value = v; }, rep.address);
  await T('검색 클릭', domClick(await pop.$('#btnImgSrch')));
  await T('검색 결과 대기', pop.waitForFunction(() => document.querySelectorAll('a#hi').length > 0 || /검색결과가 없|결과가 없/.test(document.body.innerText), { ...POLL, timeout: 15000 }).catch(() => {}));
  const picked = await T('우편번호 일치 결과 찾기', pop.evaluateHandle((zip) => {
    const as = [...document.querySelectorAll('a')].filter((e) => e.offsetParent !== null);
    const pairs = [];
    as.forEach((a, i) => { if (a.id === 'hi') { let z = ''; for (let j = i - 1; j >= 0; j--) { const t = as[j].textContent.trim(); if (/^\d{5}$/.test(t)) { z = t; break; } } pairs.push({ a, z }); } });
    const m = pairs.filter((p) => p.z === zip);
    return m.length === 1 ? m[0].a : null;
  }, rep.postcode));
  if (!(await picked.evaluate((e) => !!e).catch(() => false))) { await shot('zip-nomatch', pop); stop(`주소 검색 결과 중 우편번호 ${rep.postcode} 일치가 1건이 아님`); }
  log(`주소 선택: ${await picked.evaluate((e) => e.textContent.replace(/\s+/g, ' ').trim())}`);
  await T('주소 클릭', domClick(picked));
  await T('상세주소 칸 대기', pop.waitForFunction(() => { const e = document.getElementById('reg_address2'); return e && e.offsetParent !== null; }, { ...POLL, timeout: 10000 }));
  const detail = (rep.address_detail || '').trim();
  const apt = detail.match(/(\d+)\s*동\s*(\d+)\s*호/);
  const hasApt = await pop.$('#apart_yn1');
  if (apt && hasApt) {
    await T('아파트 방식 선택', pop.evaluate((d, h) => {
      document.getElementById('apart_yn1').click();
      document.getElementById('reg_apartdong').value = d; document.getElementById('reg_apartho').value = h;
    }, apt[1], apt[2]));
    log(`상세주소: 아파트 ${apt[1]}동 ${apt[2]}호`);
  } else {
    if (!detail) stop('TMS 상세주소가 비어 있음 → 사람 확인 필요');
    // 라디오는 숨겨져 있고 아이콘으로 그려짐 → 마우스 클릭 금지, 화면 안 직접 클릭
    await T('상세주소 방식 선택·입력', pop.evaluate((v) => {
      document.getElementById('apart_yn2').click();
      const f = document.getElementById('reg_address2'); f.value = v; f.dispatchEvent(new Event('change', { bubbles: true }));
      return document.getElementById('apart_yn2').checked;
    }, detail));
    log(`상세주소: 그대로 "${detail}"`);
  }
  await sleep(300);
  await T('주소입력 클릭', domClick(await handleByText(pop, '주소입력')));
  await T('팝업 닫힘 대기', (async () => { for (let i = 0; i < 40 && !pop.isClosed(); i++) await sleep(250); })());
  await sleep(800);
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
  await T('모바일영수증 미신청 클릭', clickEl(mr, '모바일영수증 미신청'));
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
  await T('기존받는 분 정보검색 클릭', clickText('기존받는 분 정보검색', { exact: false }));
  pop = await T('받는 분 팝업 대기', waitPop('RecSearchAddress', since));
  await T('받는 분 이력 대기', pop.waitForFunction((n) => [...document.querySelectorAll('tr')].some((t) => t.textContent.includes(n)), POLL, RECEIVER_NAME));
  const recv = await pop.evaluateHandle((n) => { const row = [...document.querySelectorAll('tr')].find((t) => t.textContent.includes(n)); return [...row.querySelectorAll('a')].find((a) => a.textContent.trim() === n) || null; }, RECEIVER_NAME);
  await T('받는 분 선택', domClick(recv));
  await sleep(2000);
  const recvName = await page.evaluate(() => [...document.querySelectorAll('input[type=text]')].filter((e) => e.offsetParent !== null && e.title === '이름').map((e) => e.value));
  if (!recvName.includes(RECEIVER_NAME)) stop('받는 분 적용 안 됨');
  log(`03 받는 분: ${RECEIVER_NAME} (기존 이력)`);
  await clickNextAfter("[...document.querySelectorAll('a')].find(e=>e.textContent.replace(/\\s+/g,'').includes('기존받는분정보검색'))");

  // ── 7) 04 물품정보 = 불러오기 → 목록 추가 ──
  since = pops.length;
  await T('물품정보 불러오기 클릭', clickText('물품정보불러오기', { exact: false }));
  pop = await T('물품 팝업 대기', waitPop('MailContList', since));
  await T('물품 목록 대기', pop.waitForFunction((n) => [...document.querySelectorAll('tr')].some((t) => t.textContent.includes(n)), POLL, ITEM_NAME));
  const item = await pop.evaluateHandle((n) => { const row = [...document.querySelectorAll('tr')].find((t) => t.textContent.includes(n)); return [...row.querySelectorAll('a')].find((a) => a.textContent.trim() === n) || null; }, ITEM_NAME);
  await T('물품 선택', domClick(item));
  await sleep(2000);
  const itemInfo = await page.evaluate(() => [...document.querySelectorAll('select')].filter((s) => s.offsetParent !== null && ['중량', '크기', '내용품코드'].includes(s.title)).map((s) => `${s.title}=${(s.options[s.selectedIndex] || {}).text}`));
  log(`04 물품: ${ITEM_NAME} / ${itemInfo.join(' / ')}`);
  if ((await listRows()).length) stop('목록 추가 전에 이미 행이 있음');
  await T('받는 분 목록에 추가', clickText('받는 분 목록에 추가'));
  await sleep(2000);
  const rows = await listRows();
  if (rows.length !== 1 || !rows[0].includes(rep.name) || !rows[0].includes(RECEIVER_NAME)) stop(`받는 분 목록이 1건이 아님/내용 불일치: ${rows.length}건`);
  log('05 목록: 정확히 1건');

  // ── 8) 05 주소검증 ──
  const before = dialogs.length;
  await T('주소검증 클릭', clickText('주소검증'));
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
  await T('결제카드검증 클릭', clickText('결제카드검증'));
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
