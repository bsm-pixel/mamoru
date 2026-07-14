/* ──────────────────────────────────────────────────────────────
   editor.js — 문구 편집기 (로컬 전용)
   카탈로그 카드/카피풀의 '문구'(라벨·이름·설명·카피)만 입력칸으로 띄워 직접 수정.
   · 입력 즉시 localStorage 임시저장 → 새로고침/서버다운/파일수정과 무관하게 입력 보존(손실 방지).
   · [저장] = 임시저장 내용을 실제 JSON 파일에 영구 기록(preview.bat 서버 필요).
   · 작업대(workspace)와 BroadcastChannel 실시간 연동.
   SVG 등은 보존(문구 칸만 노출).
   ────────────────────────────────────────────────────────────── */

const FILES = {};                       // _src -> 객체 (입력 시 갱신)
const CH = (() => { try { return new BroadcastChannel('mamoru-catalog'); } catch (e) { return null; } })();
const LS_KEY = 'mamoru-editor-edits-v1'; // 미저장 편집 임시저장
let pending = {};                        // "src|path" -> value (아직 파일에 저장 안 된 편집)
try { pending = JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch (e) { pending = {}; }

window.addEventListener('DOMContentLoaded', async () => {
  try {
    await Catalog.load();
  } catch (e) {
    document.getElementById('ed').innerHTML = '<div class="ed-loading">로드 실패: ' + e.message + '</div>';
    return;
  }
  Catalog.cards.forEach(o => { FILES[o._src] = o; });
  Catalog.copy.forEach(o => { FILES[o._src] = o; });
  applyPending();   // 임시저장된 편집 복원 (손실 방지)
  render();
  bindEditorEvents();
  bindActions();
  bindScrollMemory();   // 새로고침돼도 보던 자리 유지
  updateUnsaved();
  pingServer();     // 저장 서버 연결 확인 → 배너
});

/* 임시저장(pending) 복원 + 🧹 유령 키 청소
   ⚠️ pending 은 localStorage 라 세션을 넘어 남는다. 카탈로그 파일이 지워지거나 이름이 바뀌면
      '지금은 없는 파일' 키가 남고, 그 하나 때문에 [전체 저장]이 통째로 죽었다(saveFile 이 예외). */
function applyPending() {
  let purged = 0;
  Object.keys(pending).forEach(k => {
    const i = k.indexOf('|'); const src = k.slice(0, i), path = k.slice(i + 1);
    if (!FILES[src]) { delete pending[k]; purged++; return; }   // 유령 키 제거
    try { setByPath(FILES[src], path, pending[k]); } catch (e) {}
  });
  if (purged) { savePending(); console.warn('[editor] 없는 파일의 임시저장 ' + purged + '건 정리됨'); }
}
function savePending() { try { localStorage.setItem(LS_KEY, JSON.stringify(pending)); } catch (e) {} }
function updateUnsaved() {
  const srcs = new Set(Object.keys(pending).map(k => k.split('|')[0]));
  const el = document.getElementById('edUnsaved');
  if (el) el.textContent = srcs.size ? ('● 미저장 ' + srcs.size + '개 (자동 임시저장됨)') : '';
}
function banner(msg, cls) {
  const el = document.getElementById('edBanner');
  if (!el) return;
  el.textContent = msg; el.className = 'ed-banner ' + (cls || ''); el.style.display = msg ? '' : 'none';
}

/* 저장 서버(preview.bat 새 버전) 연결 확인 */
async function pingServer() {
  try {
    const r = await fetch('/__save', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{"file":"__ping__","content":"{}"}' });
    if (r.status === 403) banner('저장 준비됨 — [저장]/[전체 저장] 으로 파일에 기록됩니다.', 'ok');
    else if (r.status === 404) banner('⚠ 저장 서버 미연결 — preview.bat(검은 창) 닫았다 다시 켜야 [저장]이 됩니다. (입력은 자동 임시저장 중이라 안 날아갑니다)', 'warn');
    else banner('', '');
  } catch (e) {
    banner('⚠ 로컬 서버 미연결 — preview.bat 로 열어야 저장됩니다. (입력은 자동 임시저장 중)', 'warn');
  }
}

function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); }

function fld(label, src, p, val, multiline) {
  const a = 'data-src="' + esc(src) + '" data-path="' + esc(p) + '"';
  const ctl = multiline
    ? '<textarea rows="2" ' + a + '>' + esc(val) + '</textarea>'
    : '<input type="text" ' + a + ' value="' + esc(val) + '">';
  return '<label class="ed-fld"><span>' + esc(label) + '</span>' + ctl + '</label>';
}
function fileBlock(src, name, bodyHtml) {
  // data-file = 재렌더 후 같은 카드를 다시 찾기 위한 앵커 (스크롤 위치 유지용)
  return '<div class="ed-file" data-file="' + esc(src) + '"><div class="ed-file__bar">'
    + '<span class="ed-file__name">' + esc(name) + '</span>'
    + '<span class="ed-file__path">catalog/' + esc(src) + '</span>'
    + '<button class="ed-save" data-save="' + esc(src) + '">저장</button>'
    + '<span class="ed-status" data-status="' + esc(src) + '"></span>'
    + '</div><div class="ed-body">' + bodyHtml + '</div></div>';
}
function cardForm(c) {
  const src = c._src;
  let b = fld('카드 라벨 (label_ko)', src, 'label_ko', c.label_ko || '');
  if (c.label_subtitle_ko !== undefined) b += fld('카드 부제 (label_subtitle_ko)', src, 'label_subtitle_ko', c.label_subtitle_ko || '');
  (c.options || []).forEach((o, i) => {
    b += '<div class="ed-opt">'
      + '<button class="ed-del" data-del-src="' + esc(src) + '" data-del-idx="' + i + '">🗑 삭제</button>'
      + fld('표시값 (크게 표시 · 내부 식별자 겸용, 짧게)', src, 'options.' + i + '.id', o.id || '')
      + '<div class="ed-grid2">'
      + fld('이름 (name_ko)', src, 'options.' + i + '.name_ko', o.name_ko || '')
      + fld('영문 (name_en)', src, 'options.' + i + '.name_en', o.name_en || '')
      + '</div>'
      + fld('설명 (description_ko)', src, 'options.' + i + '.description_ko', o.description_ko || '', true)
      + fld('SVG 아이콘 (선택 · <svg …>…</svg> 코드 붙여넣기 · 색은 currentColor 권장)', src, 'options.' + i + '.svg_inline', o.svg_inline || '', true)
      + '</div>';
  });
  b += '<button class="ed-add" data-add-src="' + esc(src) + '">＋ 옵션 추가</button>';
  return fileBlock(src, c.card_type, b);
}
function copyForm(c) {
  const src = c._src;
  let b = fld('풀 이름 (빌더 UI용 · 페이지 미표시)', src, 'label_ko', c.label_ko || '');
  (c.options || []).forEach((o, i) => {
    b += '<div class="ed-opt"><span class="ed-opt__id">' + esc(o.id) + '</span>'
      + '<button class="ed-del" data-del-src="' + esc(src) + '" data-del-idx="' + i + '">🗑 삭제</button>'
      + fld('카피 (실제 표시 문구)', src, 'options.' + i + '.text', o.text || '', true)
      + ('tone' in o ? fld('톤 (내부 힌트 · 페이지 미표시)', src, 'options.' + i + '.tone', o.tone || '') : '')
      + '</div>';
  });
  b += '<button class="ed-add" data-add-src="' + esc(src) + '">＋ 옵션 추가</button>';
  return fileBlock(src, c.copy_type, b);
}

const COMMON_CARDS = ['handle_grip', 'handle_camel', 'grade'];
const TYPE_CARDS = ['blade_edge', 'blade_edge_long', 'blade_edge_dry', 'blade_design', 'thinning_teeth', 'thinning_holes', 'thinning_reduction'];

/* 실제로 스크롤되는 조상(없으면 문서) — 편집기는 단독/스튜디오 iframe 양쪽에서 열린다 */
function scroller(el) {
  for (let p = el; p && p !== document.body; p = p.parentElement) {
    const ov = getComputedStyle(p).overflowY;
    if ((ov === 'auto' || ov === 'scroll') && p.scrollHeight > p.clientHeight) return p;
  }
  return document.scrollingElement || document.documentElement;
}

/* anchorSrc: 방금 건드린 파일 블록. 재렌더 후 그 블록이 '화면 같은 자리'에 오도록 스크롤 보정.
   (innerHTML 통째 교체라 아무것도 안 하면 맨 위로 튄다 — 옵션 지울 때마다 위로 올라가던 원인) */
function render(anchorSrc) {
  const ed = document.getElementById('ed');
  const sc = scroller(ed);
  const before = anchorSrc ? ed.querySelector('[data-file="' + CSS.escape(anchorSrc) + '"]') : null;
  const beforeTop = before ? before.getBoundingClientRect().top : null;
  const prev = sc.scrollTop;

  const byType = ct => Catalog.byCardType[ct];
  let h = '';
  h += '<div class="ed-h2">① 공통 카드 (전 종류 공용)</div>';
  COMMON_CARDS.forEach(ct => { const c = byType(ct); if (c) h += cardForm(c); });
  h += '<div class="ed-h2">② 가위 종류별 특징 카드</div>';
  TYPE_CARDS.forEach(ct => { const c = byType(ct); if (c) h += cardForm(c); });
  const shown = new Set([...COMMON_CARDS, ...TYPE_CARDS]);
  const rest = Catalog.cards.filter(c => !shown.has(c.card_type));
  if (rest.length) { h += '<div class="ed-h2">기타 카드</div>'; rest.forEach(c => { h += cardForm(c); }); }
  h += '<div class="ed-h2">③ 섹션 문구 풀</div>';
  Catalog.copy.forEach(c => { h += copyForm(c); });
  ed.innerHTML = h;

  // 스크롤 복원 — 앵커 블록이 '화면에서 보이던 그 자리'로 돌아오게 현재 스크롤을 상대 보정.
  // ⚠️ 여기서 prev(교체 전 값)를 더하면 안 된다. 스튜디오는 #ed 자체가 스크롤 박스라
  //    innerHTML 교체 순간 scrollTop 이 0으로 리셋되고, 단독 편집기는 문서 스크롤이 유지된다.
  //    두 경우 모두 맞추려면 '지금 값' 기준으로 (새 위치 - 옛 위치)만큼만 밀어야 한다.
  if (beforeTop !== null) {
    const after = ed.querySelector('[data-file="' + CSS.escape(anchorSrc) + '"]');
    if (after) { sc.scrollTop += after.getBoundingClientRect().top - beforeTop; return; }
  }
  sc.scrollTop = prev;
}

/* 새로고침(라이브리로드 포함)돼도 보던 자리 유지 — 스크롤 위치를 세션에 기억 */
const SS_SCROLL = 'mamoru-editor-scroll-v1';
function bindScrollMemory() {
  const ed = document.getElementById('ed');
  const sc = scroller(ed);
  const saved = parseInt(sessionStorage.getItem(SS_SCROLL) || '0', 10);
  if (saved > 0) requestAnimationFrame(() => { sc.scrollTop = saved; });
  let t;
  (sc === document.scrollingElement ? window : sc).addEventListener('scroll', () => {
    clearTimeout(t);
    t = setTimeout(() => { try { sessionStorage.setItem(SS_SCROLL, String(sc.scrollTop)); } catch (e) {} }, 120);
  }, { passive: true });
}

/* 이벤트 위임 — #ed 에 1회만 바인딩 (재렌더에도 유지, 중복 방지) */
let editorBound = false;
function bindEditorEvents() {
  if (editorBound) return; editorBound = true;
  const ed = document.getElementById('ed');
  ed.addEventListener('input', e => {
    const t = e.target;
    if (!t.dataset || !t.dataset.path) return;
    setByPath(FILES[t.dataset.src], t.dataset.path, t.value);
    pending[t.dataset.src + '|' + t.dataset.path] = t.value; // 즉시 임시저장
    savePending();
    updateUnsaved();
    if (CH) CH.postMessage({ src: t.dataset.src, path: t.dataset.path, value: t.value });
  });
  ed.addEventListener('click', e => {
    const save = e.target.closest('[data-save]'); if (save) { saveFile(save.dataset.save); return; }
    const add = e.target.closest('[data-add-src]'); if (add) { addOption(add.dataset.addSrc); return; }
    const del = e.target.closest('[data-del-src]'); if (del) { delOption(del.dataset.delSrc, +del.dataset.delIdx); return; }
  });
}

/* 옵션 추가 — 새 옵션 id 입력받아 추가 후 즉시 파일 저장(구조 변경) */
function addOption(src) {
  const obj = FILES[src]; if (!obj) return;
  const isCard = !!obj.card_type;
  let id = window.prompt(isCard ? '새 옵션 ID (예: T36 · 화면에 크게 표시됨)' : '새 옵션 ID (내부용 · 영문/숫자)', '');
  if (id === null) return;               // 취소
  id = String(id).trim();
  if (!id) { alert('ID를 입력하세요.'); return; }
  obj.options = obj.options || [];
  if (obj.options.some(o => o.id === id)) { alert('이미 있는 ID입니다: ' + id); return; }
  obj.options.push(isCard ? { id: id, name_ko: '', name_en: '', description_ko: '' } : { id: id, text: '', tone: '' });
  render(src);
  saveFile(src);                          // 구조 변경 → 즉시 파일 기록 (preview.bat 서버 필요)
}
/* 옵션 삭제 — 최소 1개는 유지 */
function delOption(src, idx) {
  const obj = FILES[src]; if (!obj || !obj.options) return;
  if (obj.options.length <= 1) { alert('최소 1개 옵션은 남겨야 합니다.'); return; }
  const o = obj.options[idx]; if (!o) return;
  if (!window.confirm("옵션 '" + o.id + "' 을(를) 삭제할까요? (되돌릴 수 없음)")) return;
  obj.options.splice(idx, 1);
  render(src);
  saveFile(src);
}

function bindActions() {
  const all = document.getElementById('edSaveAll');
  if (all) all.addEventListener('click', saveAll);
}

function setByPath(obj, pathStr, val) {
  const keys = pathStr.split('.');
  let cur = obj;
  for (let i = 0; i < keys.length - 1; i++) cur = cur[keys[i]];
  cur[keys[keys.length - 1]] = val;
}
function status(src, msg, cls) {
  const el = document.querySelector('[data-status="' + CSS.escape(src) + '"]');
  if (el) { el.textContent = msg; el.className = 'ed-status ' + (cls || ''); }
}

async function saveFile(src) {
  const obj = FILES[src];
  // 🛡️ 없는 파일이면 조용히 죽지 말고 그 키만 버린다 (전체 저장이 통째로 멈추던 원인)
  if (!obj) {
    Object.keys(pending).forEach(k => { if (k.indexOf(src + '|') === 0) delete pending[k]; });
    savePending(); updateUnsaved();
    console.warn('[editor] 없는 파일이라 건너뜀:', src);
    return true;   // 저장할 게 없으니 실패로 치지 않는다
  }
  let content;
  try {
    const clean = JSON.parse(JSON.stringify(obj));
    delete clean._src;
    content = JSON.stringify(clean, null, 2) + '\n';
  } catch (e) {
    status(src, '내용이 깨져 저장 불가: ' + e.message, 'err');
    return false;
  }
  status(src, '저장 중…', '');
  try {
    const res = await fetch('/__save', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: 'projects/products/catalog/' + src, content }),
    });
    let d = {}; try { d = await res.json(); } catch (_) {}
    if (res.ok && d.ok) {
      Object.keys(pending).forEach(k => { if (k.indexOf(src + '|') === 0) delete pending[k]; });
      savePending(); updateUnsaved();
      status(src, '✓ 저장됨', 'ok');
      return true;
    }
    if (res.status === 404) status(src, 'preview.bat 재시작 필요 (임시저장은 유지됨)', 'err');
    else status(src, '실패: ' + (d.error || res.status), 'err');
  } catch (e) {
    status(src, '서버 미연결 (임시저장은 유지됨)', 'err');
  }
  return false;
}

async function saveAll() {
  const srcs = [...new Set(Object.keys(pending).map(k => k.split('|')[0]))];
  if (!srcs.length) { banner('저장할 변경이 없습니다.', 'ok'); return; }
  let ok = 0;
  // 🛡️ 파일 하나가 터져도 나머지는 저장한다 (예전엔 하나 터지면 전체가 조용히 멈췄다)
  for (const src of srcs) {
    try { if (await saveFile(src)) ok++; }
    catch (e) { console.error('[editor] 저장 실패:', src, e); status(src, '실패: ' + e.message, 'err'); }
  }
  if (ok === srcs.length) banner('✓ 전체 저장 완료 (' + ok + '개 파일)', 'ok');
  else banner('일부 저장 실패 (' + ok + '/' + srcs.length + ') — preview.bat 재시작 후 다시. 입력은 임시저장돼 안 날아갑니다.', 'warn');
}
