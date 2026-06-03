/* ──────────────────────────────────────────────────────────────
   editor.js — 문구 편집기 (로컬 전용)
   카탈로그 카드/카피풀의 '문구'(라벨·이름·설명·카피)만 입력칸으로 띄워
   사장님이 직접 수정 → [저장] 누르면 preview.bat 서버가 해당 JSON 파일에 기록.
   SVG(svg_inline) 등은 건드리지 않고 보존. 저장 후 작업대/빌더 자동 새로고침.
   ⚠ preview.bat(로컬 서버)로 열었을 때만 저장 작동 (POST /__save).
   ────────────────────────────────────────────────────────────── */

const FILES = {}; // _src -> 객체(원본 유지, 입력 시 갱신)

window.addEventListener('DOMContentLoaded', async () => {
  try {
    await Catalog.load();
  } catch (e) {
    document.getElementById('ed').innerHTML = '<div class="ed-loading">로드 실패: ' + e.message + '</div>';
    return;
  }
  Catalog.cards.forEach(o => { FILES[o._src] = o; });
  Catalog.copy.forEach(o => { FILES[o._src] = o; });
  render();
});

function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); }

/* 입력 필드 — data-src(파일) + data-path(필드 경로) */
function fld(label, src, p, val, multiline) {
  const a = 'data-src="' + esc(src) + '" data-path="' + esc(p) + '"';
  const ctl = multiline
    ? '<textarea rows="2" ' + a + '>' + esc(val) + '</textarea>'
    : '<input type="text" ' + a + ' value="' + esc(val) + '">';
  return '<label class="ed-fld"><span>' + esc(label) + '</span>' + ctl + '</label>';
}

function fileBlock(src, name, bodyHtml) {
  return '<div class="ed-file">'
    + '<div class="ed-file__bar">'
    + '<span class="ed-file__name">' + esc(name) + '</span>'
    + '<span class="ed-file__path">catalog/' + esc(src) + '</span>'
    + '<button class="ed-save" data-save="' + esc(src) + '">저장</button>'
    + '<span class="ed-status" data-status="' + esc(src) + '"></span>'
    + '</div>'
    + '<div class="ed-body">' + bodyHtml + '</div>'
    + '</div>';
}

function cardForm(c) {
  const src = c._src;
  let b = fld('카드 라벨 (label_ko)', src, 'label_ko', c.label_ko || '');
  if ('label_subtitle_ko' in c || c.label_subtitle_ko !== undefined)
    b += fld('카드 부제 (label_subtitle_ko)', src, 'label_subtitle_ko', c.label_subtitle_ko || '');
  (c.options || []).forEach((o, i) => {
    b += '<div class="ed-opt"><span class="ed-opt__id">' + esc(o.id) + '</span>'
      + '<div class="ed-grid2">'
      + fld('이름 (name_ko)', src, 'options.' + i + '.name_ko', o.name_ko || '')
      + fld('영문 (name_en)', src, 'options.' + i + '.name_en', o.name_en || '')
      + '</div>'
      + fld('설명 (description_ko)', src, 'options.' + i + '.description_ko', o.description_ko || '', true)
      + '</div>';
  });
  return fileBlock(src, c.card_type, b);
}

function copyForm(c) {
  const src = c._src;
  let b = fld('풀 라벨 (label_ko)', src, 'label_ko', c.label_ko || '');
  (c.options || []).forEach((o, i) => {
    b += '<div class="ed-opt"><span class="ed-opt__id">' + esc(o.id) + '</span>'
      + fld('카피 (text)', src, 'options.' + i + '.text', o.text || '', true)
      + ('tone' in o ? fld('톤 (tone)', src, 'options.' + i + '.tone', o.tone || '') : '')
      + '</div>';
  });
  return fileBlock(src, c.copy_type, b);
}

function render() {
  const ed = document.getElementById('ed');
  let h = '<div class="ed-h2">카드 (cards)</div>';
  Catalog.cards.forEach(c => { h += cardForm(c); });
  h += '<div class="ed-h2">문구 풀 (copy_pool)</div>';
  Catalog.copy.forEach(c => { h += copyForm(c); });
  ed.innerHTML = h;

  // 입력 → 메모리 객체 갱신
  ed.addEventListener('input', e => {
    const t = e.target;
    if (!t.dataset || !t.dataset.path) return;
    setByPath(FILES[t.dataset.src], t.dataset.path, t.value);
  });
  // 저장
  ed.addEventListener('click', e => {
    const btn = e.target.closest('[data-save]');
    if (btn) saveFile(btn.dataset.save);
  });
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
  const clean = JSON.parse(JSON.stringify(obj));
  delete clean._src; // 로더가 붙인 필드 제거
  const content = JSON.stringify(clean, null, 2) + '\n';
  status(src, '저장 중…', '');
  try {
    const res = await fetch('/__save', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file: 'projects/products/catalog/' + src, content }),
    });
    const d = await res.json();
    if (res.ok && d.ok) status(src, '✓ 저장됨 (작업대 자동 새로고침)', 'ok');
    else status(src, '실패: ' + (d.error || res.status), 'err');
  } catch (e) {
    status(src, '서버 연결 실패 — preview.bat 로 열었는지 확인', 'err');
  }
}
