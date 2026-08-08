// 청첩장 참석여부(RSVP) 응답을 로컬 파일로 내보내기
//   - Supabase(wedding_rsvp)에서 전체 응답을 가져와
//   - BSMKHJ/_rsvp_responses.html (보기용) + _rsvp_responses.csv (엑셀용) 생성
//   - 두 파일은 .gitignore 처리되어 인터넷에 배포되지 않음(로컬 전용)
// 실행: node tools/rsvp_export.mjs   (또는 루트의 RSVP_응답보기.bat 더블클릭)

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// TMS .env.local 에서 Supabase URL + 서비스키 읽기 (커밋되지 않는 파일)
const envPath = path.join(ROOT, 'projects/Total_Management_System/app/.env.local');
let env = '';
try { env = fs.readFileSync(envPath, 'utf8'); }
catch { console.error('❌ .env.local 을 못 찾았습니다:', envPath); process.exit(1); }
const pick = k => (env.match(new RegExp('^' + k + '=(.*)$', 'm')) || [])[1]?.trim();
const URL_ = pick('NEXT_PUBLIC_SUPABASE_URL');
const SVC  = pick('SUPABASE_SERVICE_ROLE_KEY');
if (!URL_ || !SVC) { console.error('❌ Supabase 키를 .env.local 에서 못 읽었습니다'); process.exit(1); }

// 조회
const res = await fetch(`${URL_}/rest/v1/wedding_rsvp?select=*&order=created_at.desc`, {
  headers: { apikey: SVC, Authorization: 'Bearer ' + SVC },
});
if (!res.ok) { console.error('❌ 조회 실패', res.status, await res.text()); process.exit(1); }
const rows = await res.json();

// 요약
const going    = rows.filter(r => r.attending === '참석');
const notGoing = rows.filter(r => r.attending === '미참석');
const headSum  = going.reduce((s, r) => s + (r.headcount || 1), 0);
const mealSum  = going.filter(r => r.meal === '예정').reduce((s, r) => s + (r.headcount || 1), 0);

const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' }[c]));
const fmt = t => { const d = new Date(t); const p = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`; };
const stamp = fmt(Date.now());

const rowsHtml = rows.length ? rows.map((r, i) => `
      <tr class="${r.attending === '참석' ? '' : 'no'}">
        <td class="c">${rows.length - i}</td>
        <td><span class="tag ${r.attending === '참석' ? 'y' : 'n'}">${esc(r.attending)}</span></td>
        <td>${esc(r.name)}</td>
        <td>${esc(r.phone) || '–'}</td>
        <td class="c">${r.headcount ?? '–'}</td>
        <td class="c">${esc(r.meal) || '–'}</td>
        <td>${esc(r.message) || ''}</td>
        <td class="dim">${fmt(r.created_at)}</td>
      </tr>`).join('') : `<tr><td colspan="8" class="empty">아직 응답이 없습니다.</td></tr>`;

const html = `<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>참석 응답 — 백성민 ♥ 김혜정</title>
<style>
:root{--ink:#2c2a27;--muted:#8b8177;--line:#e7e0d5;--gold:#b49a6a;--bg:#faf8f4}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);
  font-family:'Pretendard',system-ui,'Malgun Gothic',sans-serif;padding:28px}
h1{font-size:20px;margin:0 0 4px}.sub{color:var(--muted);font-size:13px;margin-bottom:20px}
.cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:22px}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:14px 18px;min-width:120px}
.card .n{font-size:26px;font-weight:700}.card .l{font-size:12px;color:var(--muted);margin-top:2px}
.card .n b{color:var(--gold)}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;font-size:14px}
th,td{padding:11px 12px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
th{background:#f3eee5;font-weight:600;color:var(--muted);font-size:12px;letter-spacing:.03em;white-space:nowrap}
td.c{text-align:center}.dim{color:var(--muted);font-size:12px;white-space:nowrap}
tr.no{background:#faf7f3;color:var(--muted)}
.tag{font-size:12px;font-weight:600;padding:3px 9px;border-radius:20px}
.tag.y{background:#eef5ec;color:#4a7a48}.tag.n{background:#f3ece9;color:#a06a5a}
.empty{text-align:center;color:var(--muted);padding:40px}
.bar{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px}
button{font:inherit;border:1px solid var(--gold);color:var(--gold);background:#fff;border-radius:8px;padding:8px 14px;cursor:pointer}
</style></head><body>
<h1>참석 응답 · 백성민 ♥ 김혜정</h1>
<div class="sub">생성 ${stamp} · 새로고침은 <b>RSVP_응답보기.bat</b> 다시 실행</div>
<div class="cards">
  <div class="card"><div class="n">${rows.length}</div><div class="l">총 응답</div></div>
  <div class="card"><div class="n">${going.length}<b>팀</b></div><div class="l">참석</div></div>
  <div class="card"><div class="n">${headSum}<b>명</b></div><div class="l">참석 인원(동반 포함)</div></div>
  <div class="card"><div class="n">${mealSum}<b>명</b></div><div class="l">식사 예정</div></div>
  <div class="card"><div class="n">${notGoing.length}</div><div class="l">미참석</div></div>
</div>
<div class="bar"><div></div><button onclick="dl()">CSV 다운로드</button></div>
<table>
  <thead><tr><th>#</th><th>참석</th><th>성함</th><th>연락처</th><th>인원</th><th>식사</th><th>전하고 싶은 말</th><th>접수시각</th></tr></thead>
  <tbody>${rowsHtml}</tbody>
</table>
<script>
function dl(){ location.href='_rsvp_responses.csv'; }
</script>
</body></html>`;

// CSV (엑셀 한글 대비 BOM)
const csvHead = ['번호','참석','성함','연락처','인원','식사','메시지','접수시각'];
const csvRows = rows.map((r, i) => [
  rows.length - i, r.attending, r.name, r.phone || '', r.headcount ?? '', r.meal || '',
  (r.message || '').replace(/\r?\n/g, ' '), fmt(r.created_at),
].map(v => `"${String(v).replace(/"/g, '""')}"`).join(','));
const csv = '﻿' + [csvHead.join(','), ...csvRows].join('\r\n');

const outHtml = path.join(ROOT, 'BSMKHJ/_rsvp_responses.html');
const outCsv  = path.join(ROOT, 'BSMKHJ/_rsvp_responses.csv');
fs.writeFileSync(outHtml, html, 'utf8');
fs.writeFileSync(outCsv, csv, 'utf8');

console.log(`✅ 응답 ${rows.length}건 저장 완료`);
console.log(`   참석 ${going.length}팀 / ${headSum}명 (식사 ${mealSum}명), 미참석 ${notGoing.length}`);
console.log(`   → BSMKHJ/_rsvp_responses.html (보기)`);
console.log(`   → BSMKHJ/_rsvp_responses.csv  (엑셀)`);
