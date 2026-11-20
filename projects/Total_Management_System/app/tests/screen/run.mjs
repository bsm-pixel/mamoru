// 화면 점검 (2026-11-16) — 실행: npm run check:screen   (빌드를 다시 하지 않으려면: npm run check:screen -- --skip-build)
//
// 왜 만들었나: 「주소에 건 붙이기」를 PC 로만 시험하고 배포했다가 휴대폰에서 닫은 상세가 다시 열렸다(2026-11-15).
//             npm run check(계산·흐름 점검)는 화면을 띄우지 않아 이런 문제를 못 잡는다.
//
// 하는 일: TMS 를 가짜 수파베이스(tests/screen/fake-supabase.mjs)에 붙여 따로 빌드(.next-check)하고,
//          실제 크롬으로 PC 폭 · 휴대폰 폭에서 열어 본다. 운영 자료 · 알림톡 · 택배 · 아임웹에는 닿지 않는다.
//
// 보는 것
//   PC     ① 화면마다 브라우저 탭 이름  ② 빠른 찾기(Ctrl+K)가 Esc 로 닫히고 초점이 밖으로 안 나감  ③ 계약서 유의사항 문구
//          · 카드 모양 목록의 방향키 · 숫자 칸 쉼표 표시
//   휴대폰 ④ 상세를 열고 닫으면 다시 열리지 않음(Esc · 뒤로 가기)  ⑤ 화면이 옆으로 넘치지 않음  ⑥ 화면이 깨지지 않음
import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import puppeteer from 'puppeteer-core';
import { startFakeSupabase, fakeSessionCookie } from './fake-supabase.mjs';

const APP_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const APP_PORT = 3123;
const DB_PORT = 3124;
const BASE = `http://localhost:${APP_PORT}`;
const DIST = '.next-check';
const SKIP_BUILD = process.argv.includes('--skip-build');

const CHROME = process.env.CHROME_PATH || [
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
  '/usr/bin/google-chrome',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
].find((p) => fs.existsSync(p));

// 가짜 자료 — 실제 고객이 아니다
const FIXTURES = {
  tables: {
    customers: [
      { id: 'c-check-1', name: '점검고객가', phone: '010-0000-0001', phone_normalized: '01000000001', customer_type: 'B2C', postcode: '00001', address_road: '점검시 점검로 1', address_detail: '101호', tags: [], memo: null, created_at: '2026-11-01T00:00:00Z', updated_at: '2026-11-01T00:00:00Z', total_spent: 0, purchase_count: 0 },
      { id: 'c-check-2', name: '점검고객나', phone: '010-0000-0002', phone_normalized: '01000000002', customer_type: 'B2C', tags: [], memo: null, created_at: '2026-11-02T00:00:00Z', updated_at: '2026-11-02T00:00:00Z', total_spent: 0, purchase_count: 0 },
      { id: 'c-check-3', name: '점검거래처', phone: '010-0000-0003', phone_normalized: '01000000003', customer_type: 'dealer', company_name: null, tags: [], memo: null, default_repair_price: 10000, created_at: '2026-11-03T00:00:00Z', updated_at: '2026-11-03T00:00:00Z', total_spent: 0, purchase_count: 0 },
      { id: 'c-check-4', name: '단가없는상사', phone: '010-0000-0004', phone_normalized: '01000000004', customer_type: 'dealer', company_name: null, tags: [], memo: null, default_repair_price: null, created_at: '2026-11-04T00:00:00Z', updated_at: '2026-11-04T00:00:00Z', total_spent: 0, purchase_count: 0 },
    ],
    products: [
      { id: 'p-check-1', name: '점검가위 가', sku: 'CHK-001', category: 'SCI', product_group: null, price: 1250000, price_purchase: 400000, stock_quantity: 3, is_active: true, zone_raw: 3, zone_ready: 0, zone_display: 0, created_at: '2026-11-01T00:00:00Z', updated_at: '2026-11-01T00:00:00Z' },
      { id: 'p-check-2', name: '점검가위 나', sku: 'CHK-002', category: 'SCI', product_group: null, price: 980000, price_purchase: 300000, stock_quantity: 5, is_active: true, zone_raw: 5, zone_ready: 0, zone_display: 0, created_at: '2026-11-02T00:00:00Z', updated_at: '2026-11-02T00:00:00Z' },
    ],
  },
  rpc: {},
};

// [주소, 탭 이름] — 탭 이름은 「이름 · TMS」 또는 「이름 (건수) · TMS」
const PAGES = [
  ['/dashboard', '대시보드'], ['/sales', '판매 관리'], ['/orders', '주문관리'], ['/repairs', '복원수리'],
  ['/consultations', '상담관리'], ['/customers', '고객 관리'], ['/products', '제품 관리'], ['/inventory', '재고'],
  ['/purchasing', '매입관리'], ['/suppliers', 'B2B 거래처'], ['/reviews', '리뷰관리'],
  ['/expenses', '경비 관리'], ['/cashflow', '입출금 관리'], ['/returns', '반품 · 교환'], ['/contracts', '전자 계약서'],
  ['/settings', '설정'],
];
// 메뉴를 눌러 옮겨 다닐 때(새로 고침 없이)도 이름이 남는지 볼 화면
const CLICK_THROUGH = ['/customers', '/products', '/reviews', '/suppliers', '/purchasing', '/settings', '/sales'];

const results = [];
const ok = (name) => { results.push({ name, pass: true }); console.log(`  통과  ${name}`); };
const fail = (name, why) => { results.push({ name, pass: false, why }); console.log(`  실패  ${name} — ${why}`); };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function killTree(child) {
  if (!child || child.killed) return;
  if (process.platform === 'win32') spawnSync('taskkill', ['/pid', String(child.pid), '/T', '/F'], { stdio: 'ignore' });
  else child.kill('SIGTERM');
}

async function waitForApp(timeoutMs) {
  const until = Date.now() + timeoutMs;
  while (Date.now() < until) {
    try { const r = await fetch(`${BASE}/login`, { redirect: 'manual' }); if (r.status < 500) return true; } catch { /* 아직 안 떴다 */ }
    await sleep(500);
  }
  return false;
}

const titleOk = (title, name) => title === `${name} · TMS` || (title.startsWith(`${name} (`) && title.endsWith(' · TMS'));

async function newPage(browser, { mobile }) {
  const page = await browser.newPage();
  if (mobile) {
    await page.setViewport({ width: 390, height: 844, isMobile: true, hasTouch: true, deviceScaleFactor: 2 });
    await page.setUserAgent('Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1');
  } else {
    await page.setViewport({ width: 1440, height: 900 });
  }
  await page.setCookie({ name: 'sb-127-auth-token', value: fakeSessionCookie(), url: BASE });
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e.message || e).slice(0, 200)));
  return { page, errors };
}

const open = (page, url) => page.goto(BASE + url, { waitUntil: 'networkidle2', timeout: 60000 });
const broken = (page) => page.evaluate(() => /문제가 생겼습니다|Application error/.test(document.body.innerText));

async function pcChecks(browser) {
  console.log('\n[PC 1440]');
  const { page, errors } = await newPage(browser, { mobile: false });

  // ① 탭 이름 — 주소로 바로 열었을 때 (전엔 건수가 없는 화면 15곳이 「MAMORU TMS」로 돌아갔다)
  const wrong = [];
  for (const [url, name] of PAGES) {
    await open(page, url);
    if (!page.url().includes(url)) { wrong.push(`${url}: 다른 곳으로 넘어감(${new URL(page.url()).pathname})`); continue; }
    await sleep(700);
    const t = await page.title();
    if (!titleOk(t, name)) wrong.push(`${url}: 「${t}」`);
    if (await broken(page)) wrong.push(`${url}: 화면이 깨짐`);
  }
  if (wrong.length) fail('탭 이름 — 주소로 열었을 때', wrong.join(' / ')); else ok(`탭 이름 — 주소로 열었을 때 (${PAGES.length}개 화면)`);

  // ① 탭 이름 — 왼쪽 메뉴를 눌러 옮겨 다닐 때
  await open(page, '/dashboard');
  const wrong2 = [];
  for (const url of CLICK_THROUGH) {
    const name = PAGES.find((p) => p[0] === url)[1];
    const clicked = await page.evaluate((href) => {
      const a = Array.from(document.querySelectorAll(`a[href="${href}"]`)).find((el) => el.getClientRects().length > 0);
      if (!a) return false;
      a.click();
      return true;
    }, url);
    if (!clicked) { wrong2.push(`${url}: 메뉴에서 못 찾음`); continue; }
    try { await page.waitForFunction((u) => location.pathname === u, { timeout: 30000 }, url); } catch { wrong2.push(`${url}: 이동 안 됨`); continue; }
    await sleep(2500); // 이동 뒤 한참 있다가도 이름이 그대로인지
    const t = await page.title();
    if (!titleOk(t, name)) wrong2.push(`${url}: 「${t}」`);
  }
  if (wrong2.length) fail('탭 이름 — 메뉴를 눌러 옮길 때', wrong2.join(' / ')); else ok(`탭 이름 — 메뉴를 눌러 옮길 때 (${CLICK_THROUGH.length}개 화면)`);

  // ② 빠른 찾기 — 초점이 어디에 있든 Esc 한 번, Tab 을 눌러도 밖으로 안 나감
  await open(page, '/customers');
  const hasDialog = () => page.evaluate(() => !!document.querySelector('[role=dialog][aria-label="빠른 찾기"]'));
  await page.keyboard.down('Control'); await page.keyboard.press('KeyK'); await page.keyboard.up('Control');
  await sleep(300);
  if (!(await hasDialog())) fail('빠른 찾기 — Ctrl+K 로 열림', '창이 안 열림');
  else {
    for (let i = 0; i < 4; i++) await page.keyboard.press('Tab');
    const inside = await page.evaluate(() => !!document.activeElement?.closest('[role=dialog][aria-label="빠른 찾기"]'));
    if (inside) ok('빠른 찾기 — Tab 을 눌러도 초점이 창 안에 있음'); else fail('빠른 찾기 — 초점 가두기', '초점이 뒤쪽 화면으로 나감');
    await page.evaluate(() => { const el = document.activeElement; if (el instanceof HTMLElement) el.blur(); });
    await sleep(100);
    await page.keyboard.press('Escape');
    await sleep(300);
    if (await hasDialog()) fail('빠른 찾기 — 초점이 없어도 Esc 로 닫힘', '창이 남아 있음'); else ok('빠른 찾기 — 초점이 없어도 Esc 한 번으로 닫힘');
    // 찾는 글자를 넣은 뒤에도 묻지 않고 닫혀야 한다
    await page.keyboard.down('Control'); await page.keyboard.press('KeyK'); await page.keyboard.up('Control');
    await sleep(300);
    await page.keyboard.type('zz');
    await page.keyboard.press('Escape');
    await sleep(300);
    const left = await page.evaluate(() => ({ dialog: !!document.querySelector('[role=dialog][aria-label="빠른 찾기"]'), ask: /버릴까요/.test(document.body.innerText) }));
    if (left.dialog || left.ask) fail('빠른 찾기 — 글자를 넣은 뒤 Esc', left.ask ? '"버릴까요?"를 물음' : '창이 남아 있음'); else ok('빠른 찾기 — 글자를 넣은 뒤에도 Esc 한 번으로 닫힘');
  }

  // ③ 계약서 유의사항 — 두 화면이 같은 문구
  const need = ['받으신 날부터 5일 이내에는 교환·반품이 가능합니다', '문구 각인이 진행된 제품', '5일이 지난 뒤에는 교환·환불이 어렵습니다'];
  for (const url of ['/contracts/new', '/contract']) {
    await open(page, url);
    const text = await page.evaluate(() => document.body.innerText);
    const miss = need.filter((n) => !text.includes(n));
    const old = /어떠한 명분이든/.test(text);
    if (miss.length || old) fail(`계약서 유의사항 ${url}`, old ? '옛 문구가 남아 있음' : `빠진 문장: ${miss.join(' / ')}`); else ok(`계약서 유의사항 ${url}`);
  }

  // ④-PC 카드 모양 목록에서도 ↓ 로 줄을 고른다 (제품)
  await open(page, '/products');
  await sleep(500);
  const rowCount = await page.evaluate(() => document.querySelectorAll('[data-list-row]').length);
  if (rowCount < 2) fail('제품 목록 — 방향키', `줄 표시가 ${rowCount}개뿐`);
  else {
    await page.evaluate(() => { const el = document.activeElement; if (el instanceof HTMLElement) el.blur(); });
    await page.keyboard.press('ArrowDown');
    await sleep(400);
    await page.keyboard.press('ArrowDown');
    await sleep(400);
    const sel = await page.evaluate(() => {
      const rows = Array.from(document.querySelectorAll('[data-list-row]'));
      return { idx: rows.findIndex((r) => r.getAttribute('data-selected') === 'true'), id: new URLSearchParams(location.search).get('id') };
    });
    if (sel.idx !== 1 || !sel.id) fail('제품 목록 — ↓ 두 번이면 둘째 줄', `고른 줄 ${sel.idx}, 주소 id=${sel.id}`); else ok('제품 목록(카드 모양) — ↓ 로 줄 이동');
  }

  // 금액 칸 — 치는 동안 칸 안에 쉼표가 찍히고 글자 · 음수는 안 들어간다 (제품 등록의 소매가)
  await open(page, '/products/new');
  const hasMoney = await page.evaluate(() => { const el = document.querySelector('input[inputmode=numeric]'); if (!el) return false; el.focus(); return true; });
  if (!hasMoney) fail('금액 칸 쉼표', '금액 칸을 못 찾음');
  else {
    await page.keyboard.type('-12a50000');
    await sleep(150);
    const v1 = await page.evaluate(() => document.activeElement.value);
    for (let i = 0; i < 3; i++) await page.keyboard.press('Backspace');
    await sleep(150);
    const v2 = await page.evaluate(() => document.activeElement.value);
    if (v1 !== '1,250,000' || v2 !== '1,250') fail('금액 칸 쉼표', `「${v1}」 → 지운 뒤 「${v2}」`); else ok('금액 칸 — 칸 안에 쉼표(1,250,000), 글자 · 음수는 안 들어감');
  }

  // 금액이 아닌 숫자 칸(수량 · 일수 등)에 큰 수를 치면 쉼표 붙은 값이 칸 위에 뜬다
  let numPage = null;
  for (const url of ['/products/new', '/settings?tab=repair', '/settings']) {
    await open(page, url);
    await sleep(300);
    if (await page.evaluate(() => { const el = Array.from(document.querySelectorAll('input[type=number]')).find((e) => e.getClientRects().length > 0 && !e.disabled); if (!el) return false; el.focus(); el.select(); return true; })) { numPage = url; break; }
  }
  if (!numPage) fail('숫자 칸 쉼표 표시', '숫자 칸을 못 찾음');
  else {
    await page.keyboard.type('12500');
    await sleep(200);
    const bubbles = () => page.evaluate(() => Array.from(document.querySelectorAll('body > div[aria-hidden=true]')).filter((d) => getComputedStyle(d).display !== 'none' && /^[\d,]+$/.test(d.textContent || '') && (d.textContent || '').includes(',')).map((d) => d.textContent));
    const shown = await bubbles();
    await page.evaluate(() => { const el = document.activeElement; if (el instanceof HTMLElement) el.blur(); });
    await sleep(100);
    const after = await bubbles();
    if (!shown.length) fail('숫자 칸 쉼표 표시', `${numPage}: 표시가 뜨지 않음`); else if (after.length) fail('숫자 칸 쉼표 표시', '칸을 떠나도 남아 있음'); else ok('숫자 칸(금액 아님) — 큰 수를 치면 칸 위에 쉼표 표시, 떠나면 사라짐');
  }

  // 거래처 매출 입력 — 거래처 검색에서 ↓ · Enter 로 고른다 (결과가 1건이어도)
  await open(page, '/sales/new?mode=b2b');
  await sleep(500);
  const PARTNER = 'input[placeholder="거래처 검색 (2자 이상)"]';
  if (!(await page.$(PARTNER))) fail('거래처 매출 — 거래처 검색 칸', '칸을 못 찾음');
  else {
    await page.click(PARTNER);
    await page.keyboard.type('점검');
    try { await page.waitForFunction(() => Array.from(document.querySelectorAll('button')).some((b) => /점검거래처/.test(b.textContent || '')), { timeout: 8000 }); } catch { /* 아래에서 실패로 잡힌다 */ }
    await page.keyboard.press('ArrowDown');
    await sleep(150);
    const lit = await page.evaluate(() => Array.from(document.querySelectorAll('button')).some((b) => /점검거래처/.test(b.textContent || '') && b.className.includes('bg-neutral-100')));
    await page.keyboard.press('Enter');
    await sleep(300);
    const picked = await page.evaluate((sel) => document.querySelector(sel)?.value, PARTNER);
    if (!lit) fail('거래처 매출 — ↓ 를 누르면 줄이 강조됨', '강조가 안 됨');
    else if (picked !== '점검거래처') fail('거래처 매출 — Enter 로 거래처 고르기', `칸 값 「${picked}」`);
    else ok('거래처 매출 — 거래처 검색에서 ↓ · Enter 로 고름(결과 1건)');

    // 복원수리 단가 — 단가 10,000 거래처를 골랐다가 단가 없는 거래처로 바꾸면 8,000 으로 돌아온다 (2026-11-16 앞 거래처 단가가 남던 문제)
    const repairPrice = () => page.evaluate(() => {
      const label = Array.from(document.querySelectorAll('label')).find((l) => (l.textContent || '').trim() === '단가 (원)');
      return label?.parentElement?.querySelector('input')?.value ?? null;
    });
    const first = await repairPrice();
    await page.click(PARTNER);
    await page.keyboard.down('Control'); await page.keyboard.press('KeyA'); await page.keyboard.up('Control');
    await page.keyboard.type('단가없는');
    try { await page.waitForFunction(() => Array.from(document.querySelectorAll('button')).some((b) => /단가없는상사/.test(b.textContent || '')), { timeout: 8000 }); } catch { /* 아래에서 실패로 잡힌다 */ }
    await page.keyboard.press('Enter');
    await sleep(300);
    const picked2 = await page.evaluate((sel) => document.querySelector(sel)?.value, PARTNER);
    const second = await repairPrice();
    if (first !== '10,000') fail('거래처 매출 — 거래처 기본 단가 자동 적용', `단가 칸 「${first}」`);
    else if (picked2 !== '단가없는상사') fail('거래처 매출 — 거래처 바꾸기', `칸 값 「${picked2}」`);
    else if (second !== '8,000') fail('거래처 매출 — 단가 없는 거래처로 바꾸면 8,000', `단가 칸 「${second}」`);
    else ok('거래처 매출 — 복원수리 단가: 10,000 거래처 → 단가 없는 거래처로 바꾸면 8,000');
  }

  // 전체 출고 준비표 — 버튼은 머리줄에 하나, 어느 화면에서나 열린다
  const dup = [];
  for (const url of ['/orders', '/sales', '/repairs']) {
    await open(page, url);
    await sleep(300);
    const c = await page.evaluate(() => ({
      inMain: Array.from(document.querySelectorAll('main button')).filter((b) => !b.closest('header') && /전체 출고 준비표/.test(b.textContent || '')).length,
      top: document.querySelectorAll('button[aria-label="전체 출고 준비표"]').length,
    }));
    if (c.inMain > 0) dup.push(`${url}: 화면 안에 버튼이 ${c.inMain}개 남음`);
    if (c.top !== 1) dup.push(`${url}: 머리줄 버튼 ${c.top}개`);
  }
  await open(page, '/customers');
  await page.click('button[aria-label="전체 출고 준비표"]');
  await sleep(500);
  const prepOpen = await page.evaluate(() => /출고 준비표/.test(Array.from(document.querySelectorAll('.fixed')).map((e) => e.textContent).join(' ')));
  if (!prepOpen) dup.push('고객 화면에서 눌러도 창이 안 열림');
  if (dup.length) fail('전체 출고 준비표 — 머리줄 한 곳', dup.join(' / ')); else ok('전체 출고 준비표 — 버튼은 머리줄 하나, 다른 화면(고객)에서도 열림');
  await page.keyboard.press('Escape');

  // 고객 검색 — 상세 주소에 적힌 말로도 찾고, 두 글자부터 찾는다 (2026-11-20)
  await open(page, '/customers');
  {
    const bad = [];
    const rowsText = () => page.evaluate(() => document.querySelector('main')?.textContent || '');
    const box = await page.$('main input[data-list-search]');
    await box.click();
    await page.keyboard.type('점');
    await sleep(900);
    let t = await rowsText();
    if (!/두 글자 이상 입력하면 찾습니다/.test(t)) bad.push('한 글자일 때 안내가 안 보임');
    if (!/점검고객나/.test(t)) bad.push('한 글자인데 목록이 걸러짐');
    await page.keyboard.down('Control'); await page.keyboard.press('KeyA'); await page.keyboard.up('Control');
    await page.keyboard.type('101호');
    try { await page.waitForFunction(() => { const x = document.querySelector('main')?.textContent || ''; return /점검고객가/.test(x) && !/점검고객나/.test(x); }, { timeout: 8000 }); } catch { bad.push('상세 주소(101호)로 찾았는데 그 고객만 남지 않음'); }
    t = await rowsText();
    if (!/주소\s*101호/.test(t)) bad.push('주소에서 걸린 까닭(주소 101호)이 목록에 안 보임');
    if (bad.length) fail('고객 검색 — 상세 주소 · 두 글자부터', bad.join(' / ')); else ok('고객 검색 — 상세 주소에 적힌 말로 찾음, 한 글자는 안내만');
  }

  // 빠른 송장 — 받는 곳을 이번 송장에서만 바꿔 보낸다 (2026-11-18). 발급 요청은 가로채서 내용만 본다(택배사에 닿지 않는다)
  await open(page, '/manual-invoices');
  await sleep(500);
  {
    const bad = [];
    let sent = null;
    await page.setRequestInterception(true);
    const onReq = (req) => {
      if (req.method() === 'POST' && req.url().endsWith('/api/manual-invoices')) {
        sent = JSON.parse(req.postData() || '{}');
        req.respond({ status: 200, contentType: 'application/json', body: JSON.stringify({ success: true, invoice: { id: 'mi-1', invoice_number: '000000000000', customer_name: sent.receiver?.name || '', customer_phone: sent.receiver?.phone || '', receiver_postcode: sent.receiver?.postcode || '', receiver_address_road: sent.receiver?.address_road || '', receiver_address_detail: sent.receiver?.address_detail || '', goods_name: sent.goods_name } }) });
      } else req.continue();
    };
    page.on('request', onReq);
    const clickText = (t) => page.evaluate((text) => { const b = Array.from(document.querySelectorAll('main button')).find((e) => (e.textContent || '').trim() === text); b?.click(); return !!b; }, t);
    const setVal = async (label, v) => { const sel = `main input[aria-label="${label}"]`; await page.click(sel); await page.keyboard.down('Control'); await page.keyboard.press('KeyA'); await page.keyboard.up('Control'); await page.keyboard.type(v); };
    try {
      const search = await page.$('main input[type="text"]');
      await search.click();
      await page.keyboard.type('점검고객가');
      await page.waitForFunction(() => Array.from(document.querySelectorAll('main button, main li, main [role="option"]')).some((b) => /점검고객가/.test(b.textContent || '')), { timeout: 8000 });
      await page.evaluate(() => { const el = Array.from(document.querySelectorAll('main button, main li, main [role="option"]')).find((b) => /점검고객가/.test(b.textContent || '')); el?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); el?.click(); });
      await sleep(400);
      const shown = await page.evaluate(() => document.querySelector('main')?.textContent || '');
      if (!/점검시 점검로 1/.test(shown)) bad.push('고른 뒤 저장된 주소가 안 보임');
      if (!(await clickText('받는 곳 바꾸기'))) bad.push('[받는 곳 바꾸기] 버튼이 없음');
      else {
        await sleep(200);
        await setVal('받는 사람', '대신받는분');
        await setVal('받는 사람 연락처', '010-9999-0000');
        await setVal('상세 주소', '909호');
        await sleep(150);
        const note = await page.evaluate(() => document.querySelector('main')?.textContent || '');
        if (!/이 주소를 고객 정보에도 저장/.test(note)) bad.push('고객 정보에도 저장 선택이 안 보임');
        await clickText('확인');
        await sleep(200);
        const sum = await page.evaluate(() => document.querySelector('main')?.textContent || '');
        if (!/대신받는분/.test(sum) || !/이번 송장만 이 정보로 나갑니다/.test(sum)) bad.push('바꾼 내용이 요약에 안 보임');
        await page.click('main input[placeholder^="예: 샘플"]');
        await page.keyboard.type('점검 품목');
        await clickText('ALPS 송장 발급');
        await sleep(600);
        if (!sent) bad.push('발급 요청이 나가지 않음');
        else {
          const r = sent.receiver || {};
          if (sent.customer_id !== 'c-check-1' || r.name !== '대신받는분' || r.phone !== '010-9999-0000' || r.postcode !== '00001' || r.address_road !== '점검시 점검로 1' || r.address_detail !== '909호') bad.push(`보낸 내용이 다름 ${JSON.stringify(sent)}`);
          if (sent.save_address_to_customer !== false) bad.push('고르지 않았는데 고객 정보 저장이 켜짐');
        }
      }
    } catch (e) { bad.push(String(e.message || e).slice(0, 150)); }
    page.off('request', onReq);
    await page.setRequestInterception(false);
    if (bad.length) fail('빠른 송장 — 받는 곳 바꾸기', bad.join(' / ')); else ok('빠른 송장 — 받는 곳(이름 · 연락처 · 상세 주소)을 바꿔 발급, 고객 정보는 그대로');
  }

  // 일정 — 메뉴에서 빠졌고(2026-11-16), 등록 버튼 · 다가오는 일정은 대시보드 달력 오른쪽 칸에 있다. 옛 주소는 대시보드로 간다
  await open(page, '/schedule');
  await sleep(500);
  const sch = await page.evaluate(() => {
    const btn = (t) => Array.from(document.querySelectorAll('main button')).filter((b) => (b.textContent || '').trim() === t).length;
    return {
      path: location.pathname,
      menu: Array.from(document.querySelectorAll('a[href="/schedule"]')).length,
      consult: btn('상담 일정 등록'), repair: btn('복원수리 방문 예약'),
      upcoming: /다가오는 일정/.test(document.querySelector('main')?.textContent || ''),
    };
  });
  const schBad = [];
  if (sch.path !== '/dashboard') schBad.push(`옛 주소가 ${sch.path} 에 머묾`);
  if (sch.menu) schBad.push('메뉴에 일정이 남음');
  if (sch.consult !== 1 || sch.repair !== 1) schBad.push(`등록 버튼 ${sch.consult}/${sch.repair}개`);
  if (!sch.upcoming) schBad.push('다가오는 일정이 없음');
  if (!schBad.length) {
    await page.evaluate(() => Array.from(document.querySelectorAll('main button')).find((b) => (b.textContent || '').trim() === '상담 일정 등록')?.click());
    await sleep(500);
    const opened = await page.evaluate(() => /일정 수동 등록/.test(document.body.textContent || ''));
    if (!opened) schBad.push('상담 일정 등록 창이 안 열림');
    await page.keyboard.press('Escape');
    await sleep(300);
  }
  if (schBad.length) fail('일정 — 대시보드로 합침', schBad.join(' / ')); else ok('일정 — 메뉴에서 빠짐, 옛 주소는 대시보드로, 등록 버튼 · 다가오는 일정은 대시보드에');

  if (errors.length) fail('PC — 화면 오류 없음', [...new Set(errors)].slice(0, 5).join(' / ')); else ok('PC — 화면 오류 없음');
  await page.close();
}

async function mobileChecks(browser) {
  console.log('\n[휴대폰 390]');
  const { page, errors } = await newPage(browser, { mobile: true });

  // ④ 상세 열기 → 닫기 → 다시 열리지 않음
  await open(page, '/customers');
  const tapRow = async (name) => {
    const pt = await page.evaluate((n) => {
      const el = Array.from(document.querySelectorAll('span, p, div, td')).find((e) => e.children.length === 0 && e.textContent.trim() === n && e.getClientRects().length > 0);
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
    }, name);
    if (!pt) return false;
    await page.touchscreen.tap(pt.x, pt.y);
    return true;
  };
  const state = () => page.evaluate(() => ({ path: location.pathname, id: new URLSearchParams(location.search).get('id') }));
  // 옆에서 밀려 나오는 창이 열려 있는가 — 화면 안에 들어와 있는 고정 창
  const panelOpen = () => page.evaluate(() => Array.from(document.querySelectorAll('div.fixed.right-0.top-0')).some((el) => el.classList.contains('translate-x-0')));

  if (!(await tapRow('점검고객가'))) fail('휴대폰 — 고객 목록이 보임', '가짜 고객 줄을 못 찾음');
  else {
    await sleep(600);
    let s = await state();
    if (s.id !== 'c-check-1' || !(await panelOpen())) fail('휴대폰 — 줄을 누르면 상세가 열림', `주소 id=${s.id}, 창 열림=${await panelOpen()}`);
    else {
      ok('휴대폰 — 줄을 누르면 상세가 열리고 주소에 건이 붙음');
      await page.keyboard.press('Escape');
      await sleep(1200);
      s = await state();
      if (s.id || (await panelOpen())) fail('휴대폰 — 닫은 상세가 다시 열리지 않음(Esc)', `주소 id=${s.id}, 창 열림=${await panelOpen()}`);
      else ok('휴대폰 — 닫은 상세가 다시 열리지 않음(Esc)');

      await tapRow('점검고객나');
      await sleep(600);
      s = await state();
      if (s.id !== 'c-check-2') fail('휴대폰 — 닫은 뒤 다른 줄 열기', `주소 id=${s.id}`);
      else {
        await page.goBack({ waitUntil: 'networkidle2' }).catch(() => {});
        await sleep(1200);
        s = await state();
        if (s.path !== '/customers' || s.id || (await panelOpen())) fail('휴대폰 — 뒤로 가기 = 상세만 닫힘', `주소 ${s.path} id=${s.id}, 창 열림=${await panelOpen()}`);
        else ok('휴대폰 — 뒤로 가기 = 상세만 닫히고 목록에 남음');
      }
    }
  }

  // ⑤ 옆으로 넘침 · ⑥ 깨짐
  const over = [];
  for (const [url] of PAGES) {
    await open(page, url);
    await sleep(400);
    const m = await page.evaluate(() => ({ w: document.documentElement.scrollWidth, vw: window.innerWidth }));
    if (m.w > m.vw + 1) over.push(`${url}: ${m.w - m.vw}px 넘침`);
    if (await broken(page)) over.push(`${url}: 화면이 깨짐`);
  }
  if (over.length) fail('휴대폰 — 옆으로 넘치지 않음', over.join(' / ')); else ok(`휴대폰 — 옆으로 넘치지 않음 (${PAGES.length}개 화면)`);

  if (errors.length) fail('휴대폰 — 화면 오류 없음', [...new Set(errors)].slice(0, 5).join(' / ')); else ok('휴대폰 — 화면 오류 없음');
  await page.close();
}

async function main() {
  if (!CHROME) { console.error('크롬을 찾지 못했습니다. CHROME_PATH 로 위치를 알려 주세요.'); process.exit(2); }
  const db = await startFakeSupabase(FIXTURES, DB_PORT);

  // 운영 값은 전부 가짜로 덮는다 — .env.local 보다 여기 값이 먼저다
  const env = {
    ...process.env,
    TMS_CHECK_DIST: DIST,
    NEXT_PUBLIC_SUPABASE_URL: db.url,
    NEXT_PUBLIC_SUPABASE_ANON_KEY: 'check-anon',
    SUPABASE_SERVICE_ROLE_KEY: 'check-service',
    IMWEB_API_KEY: '', IMWEB_API_SECRET: '',
    LOTTE_API_URL: `${db.url}/none`, LOTTE_CANCEL_API_URL: `${db.url}/none`, LOTTE_TRACK_API_URL: `${db.url}/none`, LOTTE_CLIENT_KEY: '',
    KAKAO_REST_API_KEY: '', MAKE_WEBHOOK_URL: '', MAKE_REPAIR_WEBHOOK_URL: '',
    CRON_SECRET: 'check', BANK_WEBHOOK_SECRET: 'check',
    NEXT_TELEMETRY_DISABLED: '1',
  };
  const next = path.join(APP_DIR, 'node_modules', 'next', 'dist', 'bin', 'next');

  let server;
  let browser;
  let code = 1;
  try {
    if (!SKIP_BUILD || !fs.existsSync(path.join(APP_DIR, DIST, 'BUILD_ID'))) {
      console.log('점검용 빌드를 만드는 중… (1~2분)');
      const b = spawnSync(process.execPath, [next, 'build'], { cwd: APP_DIR, env, encoding: 'utf-8' });
      if (b.status !== 0) { console.error((b.stdout || '').slice(-3000)); console.error((b.stderr || '').slice(-3000)); throw new Error('점검용 빌드 실패'); }
    }
    server = spawn(process.execPath, [next, 'start', '-p', String(APP_PORT)], { cwd: APP_DIR, env, stdio: 'ignore' });
    if (!(await waitForApp(60000))) throw new Error('점검용 서버가 뜨지 않았습니다');

    browser = await puppeteer.launch({ executablePath: CHROME, headless: true, args: ['--no-sandbox'] });
    await pcChecks(browser);
    await mobileChecks(browser);

    const failed = results.filter((r) => !r.pass);
    console.log(`\n화면 점검: ${results.length - failed.length}개 통과 · ${failed.length}개 실패`);
    if (db.writes.length) console.log(`(참고) 점검 중 쓰기 요청 ${db.writes.length}건 — 가짜 서버가 받기만 함: ${[...new Set(db.writes)].join(', ')}`);
    code = failed.length ? 1 : 0;
  } catch (e) {
    console.error('화면 점검을 끝내지 못했습니다:', e.message || e);
    code = 2;
  } finally {
    if (browser) await browser.close().catch(() => {});
    killTree(server);
    await db.close();
  }
  process.exit(code);
}

main();
