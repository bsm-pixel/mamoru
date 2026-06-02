/* ──────────────────────────────────────────────────────────────
   builder.js — 빌더 컨트롤러
   · 카탈로그 로드 → 입력 패널 자동 생성(카탈로그 기반) → 실시간 미리보기
   · spec 다운로드 / 불러오기 / 완성 HTML 복사 / PC·모바일 토글
   ────────────────────────────────────────────────────────────── */

const TYPES = [
  { id: 'blunt', label: '블런트' },
  { id: 'thinning', label: '틴닝' },
  { id: 'long', label: '장가위' },
  { id: 'dry', label: '드라이' }
];

/* 작업 중인 모델 spec (단일 상태) */
let spec = defaultSpec('blunt');

function defaultSpec(type) {
  return {
    model: '', type, category_label: '', size_inch: '', weight_g: '', price_grade: 'A',
    selections: {}, copy_selections: {}, custom_fields: {}, spec_meta: {},
    lineup: [], images_folder: '', images: {}
  };
}

/* ─── 진입 ─── */
window.addEventListener('DOMContentLoaded', async () => {
  try {
    await Catalog.load();
  } catch (e) {
    document.getElementById('panel').innerHTML =
      `<div class="loaderr">카탈로그 로드 실패: ${e.message}<br><small>로컬에서 열었다면 정적 서버로 실행하세요 (file:// 는 fetch 차단).</small></div>`;
    return;
  }
  bindTopbar();
  renderPanel();
  initPreview();
});

/* ════════════ 입력 패널 ════════════
   좌측 입력 순서 = 출력 페이지 섹션 순서와 1:1 (IA 일치)
   배지 번호 = 페이지 eyebrow(— NN / TITLE) 와 매칭 */
function renderPanel() {
  const p = document.getElementById('panel');
  p.innerHTML = '';

  // 종류 (마스터 스위치)
  p.appendChild(section('종류', radioRow(TYPES.map(t => ({ id: t.id, label: t.label })),
    spec.type, v => { switchType(v); }), '·'));

  // 01 HERO — 기본 정보 + Hero 카피
  const hero = document.createElement('div');
  hero.appendChild(textField('모델명', spec.model, 'A2-55FS', v => { spec.model = v; updatePreview(); }));
  hero.appendChild(textField('카테고리 라벨 (Hero 상단)', spec.category_label, 'BLUNT 5.5 INCH', v => { spec.category_label = v; updatePreview(); }));
  const grid2 = document.createElement('div'); grid2.className = 'grid2';
  grid2.appendChild(numberField('길이 (inch)', spec.size_inch, '5.5', v => { spec.size_inch = v; updatePreview(); }));
  grid2.appendChild(numberField('무게 (g)', spec.weight_g, '58', v => { spec.weight_g = v; updatePreview(); }));
  hero.appendChild(grid2);
  hero.appendChild(textField('이미지 폴더명', spec.images_folder, spec.model || '모델명과 동일', v => { spec.images_folder = v; updatePreview(); }));
  hero.appendChild(subhead('Hero 카피 (감성 한 줄)'));
  hero.appendChild(copyChooserSingle('hero_subtitle', 'hero_subtitle_text'));
  p.appendChild(section('HERO', hero, '01'));

  // 03 ABOUT — 특성 본문
  const about = document.createElement('div');
  about.appendChild(copyChooserSingleInline('about_body', 'about_body_text', '본문 한 단락 (특성 설명)'));
  about.appendChild(copyChooserSingleInline('about_quote', 'about_quote_text', '브랜드 한 줄 (인용구)'));
  p.appendChild(section('ABOUT', about, '03'));

  // 04 PROFILE — 특성 카드 + GRIP SIZE + WEIGHT
  const prof = document.createElement('div');
  const cards = Catalog.cardsForType(spec.type).filter(c => c.card_type !== 'grade');
  if (cards.length) {
    prof.appendChild(subhead('가위 특성 카드'));
    for (const card of cards) prof.appendChild(cardSelector(card));
  }
  prof.appendChild(subhead('Grip Size (손가락 구멍)'));
  const gg = document.createElement('div'); gg.className = 'grid2';
  gg.appendChild(textField('엄지부', spec.custom_fields.grip_thumb || '', '15 × 21', v => { spec.custom_fields.grip_thumb = v; updatePreview(); }));
  gg.appendChild(textField('약지부', spec.custom_fields.grip_ring || '', '15 × 18', v => { spec.custom_fields.grip_ring = v; updatePreview(); }));
  prof.appendChild(gg);
  prof.appendChild(copyChooserSingleInline('handle_description', 'handle_description', '핸들 특성 설명'));
  prof.appendChild(subhead('Weight (무게감)'));
  prof.appendChild(selectField('무게 밴드', spec.custom_fields.weight_band || '중간',
    ['가벼움', '중간', '무거움'], v => { spec.custom_fields.weight_band = v; updatePreview(); }));
  prof.appendChild(textArea('무게 설명', spec.custom_fields.weight_description || '', '짧은 날 + 58g — 한 손 작업에 적정...', v => { spec.custom_fields.weight_description = v; updatePreview(); }));
  p.appendChild(section('PROFILE', prof, '04'));

  // 05 FOR YOU
  const fy = document.createElement('div');
  fy.appendChild(copyChooserMulti('for_you_match', '이런 분에게 맞습니다'));
  fy.appendChild(copyChooserMulti('for_you_miss', '맞지 않을 수 있습니다'));
  p.appendChild(section('FOR YOU', fy, '05'));

  // 06 SPEC
  const sm = document.createElement('div');
  sm.appendChild(radioRow(Catalog.byCardType['grade'] ? Catalog.byCardType['grade'].options.map(o => ({ id: o.id, label: `${o.id} ${o.name_en || ''}` })) : [],
    spec.price_grade, v => { spec.price_grade = v; autoFillMeta(); updatePreview(); }, '등급 (price_grade)'));
  sm.appendChild(textField('소재', spec.spec_meta.material || '', '440C JAPAN', v => { spec.spec_meta.material = v; updatePreview(); }));
  sm.appendChild(textField('베어링', spec.spec_meta.bearing || '', '베어링 장착', v => { spec.spec_meta.bearing = v; updatePreview(); }));
  const auto = document.createElement('div'); auto.className = 'hint';
  auto.innerHTML = '아래 3개는 선택 카드에서 <b>자동 채움</b> (수정 가능)';
  sm.appendChild(auto);
  sm.appendChild(textField('핸들 라벨', spec.spec_meta.handle_label || '', '카멜 + 세미오프셋', v => { spec.spec_meta.handle_label = v; updatePreview(); }, 'meta_handle_label'));
  sm.appendChild(textField('날선 라벨', spec.spec_meta.edge_label || '', 'F · 포스 (직선형)', v => { spec.spec_meta.edge_label = v; updatePreview(); }, 'meta_edge_label'));
  sm.appendChild(textField('날등 라벨', spec.spec_meta.design_label || '', 'S · 스워드 (검형)', v => { spec.spec_meta.design_label = v; updatePreview(); }, 'meta_design_label'));
  p.appendChild(section('SPEC', sm, '06'));

  // 07 SAME HANDLE
  const lu = document.createElement('div');
  lu.appendChild(textField('라인업 제목', spec.lineup_title || '', 'A2 시리즈', v => { spec.lineup_title = v; updatePreview(); }));
  lu.appendChild(textArea('라인업 모델 (한 줄에 하나)', (spec.lineup || []).map(x => typeof x === 'string' ? x : x.model).join('\n'),
    'A2-45FS\nA2-65FS\nA2-55FC', v => { spec.lineup = v.split('\n').map(s => s.trim()).filter(Boolean); updatePreview(); }));
  p.appendChild(section('SAME HANDLE', lu, '07'));

  // 13 CTA
  const cta = document.createElement('div');
  cta.appendChild(textField('맞춤 컨설팅 링크', spec.custom_fields.cta_link || '', 'https://...', v => { spec.custom_fields.cta_link = v; updatePreview(); }));
  p.appendChild(section('CTA', cta, '13'));
}

/* 종류 전환 — 기본 정보/카피는 유지, 카드 선택만 초기화 */
function switchType(type) {
  spec.type = type;
  spec.selections = {};
  // 종류 전용 카피는 풀이 다르므로 single 선택 초기화 (공용은 유지)
  delete spec.copy_selections.hero_subtitle;
  delete spec.copy_selections.about_body;
  delete spec.copy_selections.for_you_match;
  delete spec.copy_selections.for_you_miss;
  renderPanel();
  updatePreview();
}

/* 선택 카드 → 변경 시 spec.selections + 자동 메타 */
function cardSelector(card) {
  const sel = spec.selections[card.card_type];
  const opts = (card.options || []).map(o => ({
    id: o.id,
    label: `${o.id} · ${o.name_ko}${o.name_en ? ' (' + o.name_en + ')' : ''}`
  }));
  return radioRow(opts, sel, v => {
    spec.selections[card.card_type] = v;
    autoFillMeta();
    updatePreview();
  }, `${card.label_ko} — ${card.label_subtitle_ko || ''}`);
}

/* 선택값 → spec_meta 라벨 자동 채움 (사용자가 직접 수정 안 했을 때만 덮어씀) */
function autoFillMeta() {
  const edge = Catalog.cardOption('blade_edge', spec.selections.blade_edge);
  const design = Catalog.cardOption('blade_design', spec.selections.blade_design);
  const grip = Catalog.cardOption('handle_grip', spec.selections.handle_grip);
  const camel = Catalog.cardOption('handle_camel', spec.selections.handle_camel);
  if (edge) setMeta('edge_label', `${edge.id} · ${edge.name_en || ''} (${edge.name_ko})`, 'meta_edge_label');
  if (design) setMeta('design_label', `${design.id} · ${design.name_en || ''} (${design.name_ko})`, 'meta_design_label');
  if (grip || camel) {
    const parts = [];
    if (camel) parts.push(camel.name_ko);
    if (grip) parts.push(grip.name_ko);
    setMeta('handle_label', parts.join(' + '), 'meta_handle_label');
  }
}
function setMeta(key, val, inputId) {
  spec.spec_meta[key] = val;
  const inp = document.getElementById(inputId);
  if (inp) inp.value = val;
}

/* ════════════ 컨트롤 빌더 ════════════ */
function section(title, body, badge) {
  const s = document.createElement('section');
  const h = document.createElement('h3');
  if (badge) { const b = document.createElement('span'); b.className = 'sbadge'; b.textContent = badge; h.appendChild(b); }
  h.appendChild(document.createTextNode(title));
  s.appendChild(h); s.appendChild(body);
  return s;
}
function subhead(text) {
  const d = document.createElement('div'); d.className = 'subhead'; d.textContent = text;
  return d;
}
function field(label) {
  const f = document.createElement('label'); f.className = 'fld';
  const l = document.createElement('span'); l.className = 'lbl'; l.textContent = label;
  f.appendChild(l); return f;
}
function textField(label, val, ph, onInput, id) {
  const f = field(label);
  const i = document.createElement('input'); i.type = 'text'; i.value = val || ''; i.placeholder = ph || '';
  if (id) i.id = id;
  i.addEventListener('input', () => onInput(i.value));
  f.appendChild(i); return f;
}
function numberField(label, val, ph, onInput) {
  const f = field(label);
  const i = document.createElement('input'); i.type = 'text'; i.inputMode = 'decimal'; i.value = val || ''; i.placeholder = ph || '';
  i.addEventListener('input', () => onInput(i.value));
  f.appendChild(i); return f;
}
function textArea(label, val, ph, onInput) {
  const f = field(label);
  const t = document.createElement('textarea'); t.value = val || ''; t.placeholder = ph || ''; t.rows = 3;
  t.addEventListener('input', () => onInput(t.value));
  f.appendChild(t); return f;
}
function selectField(label, val, opts, onChange) {
  const f = field(label);
  const s = document.createElement('select');
  for (const o of opts) { const op = document.createElement('option'); op.value = o; op.textContent = o; if (o === val) op.selected = true; s.appendChild(op); }
  s.addEventListener('change', () => onChange(s.value));
  f.appendChild(s); return f;
}
function radioRow(opts, val, onChange, label) {
  const wrap = document.createElement('div'); wrap.className = 'fld';
  if (label) { const l = document.createElement('span'); l.className = 'lbl'; l.textContent = label; wrap.appendChild(l); }
  const row = document.createElement('div'); row.className = 'chips';
  for (const o of opts) {
    const b = document.createElement('button'); b.type = 'button'; b.className = 'chip' + (o.id === val ? ' on' : '');
    b.textContent = o.label;
    b.addEventListener('click', () => {
      row.querySelectorAll('.chip').forEach(c => c.classList.remove('on'));
      b.classList.add('on'); onChange(o.id);
    });
    row.appendChild(b);
  }
  wrap.appendChild(row); return wrap;
}

/* 카피 풀 단일 선택 (칩) + custom 입력 */
function copyChooserSingle(copyType, customKey) {
  const pool = Catalog.copyPool(copyType, spec.type);
  const wrap = document.createElement('div');
  if (!pool) { const e = document.createElement('div'); e.className = 'hint'; e.textContent = `(${copyType} 풀 없음 — 직접 입력)`; wrap.appendChild(e); }
  const cur = spec.copy_selections[copyType];
  if (pool) {
    const row = document.createElement('div'); row.className = 'chips col';
    for (const o of pool.options) {
      const b = document.createElement('button'); b.type = 'button'; b.className = 'chip' + (o.id === cur && !spec.custom_fields[customKey] ? ' on' : '');
      b.textContent = (o.text || '').replace(/\n/g, ' ') + (o.tone ? `  · ${o.tone}` : '');
      b.addEventListener('click', () => {
        spec.copy_selections[copyType] = o.id; delete spec.custom_fields[customKey];
        row.querySelectorAll('.chip').forEach(c => c.classList.remove('on')); b.classList.add('on');
        const ci = wrap.querySelector('textarea'); if (ci) ci.value = '';
        updatePreview();
      });
      row.appendChild(b);
    }
    wrap.appendChild(row);
  }
  const ta = document.createElement('textarea'); ta.rows = 2; ta.placeholder = '또는 직접 입력 (줄바꿈 가능)';
  ta.value = spec.custom_fields[customKey] || '';
  ta.addEventListener('input', () => { spec.custom_fields[customKey] = ta.value; updatePreview(); });
  wrap.appendChild(ta);
  return wrap;
}
/* 라벨 붙은 버전 */
function copyChooserSingleInline(copyType, customKey, label) {
  const f = document.createElement('div'); f.className = 'fld';
  const l = document.createElement('span'); l.className = 'lbl'; l.textContent = label; f.appendChild(l);
  f.appendChild(copyChooserSingle(copyType, customKey));
  return f;
}

/* 카피 풀 다중 선택 (체크 칩) */
function copyChooserMulti(copyType, label) {
  const pool = Catalog.copyPool(copyType, spec.type);
  const f = document.createElement('div'); f.className = 'fld';
  const l = document.createElement('span'); l.className = 'lbl';
  l.textContent = label + (pool && pool.recommend_count ? `  (권장 ${pool.recommend_count}개)` : '');
  f.appendChild(l);
  if (!pool) { const e = document.createElement('div'); e.className = 'hint'; e.textContent = `(${copyType} 풀 없음)`; f.appendChild(e); return f; }
  const chosen = new Set(spec.copy_selections[copyType] || []);
  const row = document.createElement('div'); row.className = 'chips col';
  for (const o of pool.options) {
    const b = document.createElement('button'); b.type = 'button'; b.className = 'chip' + (chosen.has(o.id) ? ' on' : '');
    b.textContent = o.text;
    b.addEventListener('click', () => {
      if (chosen.has(o.id)) { chosen.delete(o.id); b.classList.remove('on'); }
      else { chosen.add(o.id); b.classList.add('on'); }
      spec.copy_selections[copyType] = [...chosen];
      updatePreview();
    });
    row.appendChild(b);
  }
  f.appendChild(row); return f;
}

/* ════════════ 미리보기 ════════════
   iframe 골격을 한 번만 깔고(initPreview), 이후엔 내부 #root.innerHTML 만 교체.
   → srcdoc 재로딩 X → 스크롤 위치 유지 + 폰트 깜빡임 없음 */
let previewReady = false, previewTimer = null;

function initPreview() {
  const f = document.getElementById('preview');
  f.addEventListener('load', () => { previewReady = true; paintPreview(); }, { once: true });
  f.srcdoc = `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700;800&family=Outfit:wght@700;800;900&family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<link href="https://cdn.jsdelivr.net/gh/fonts-archive/Paperlogy/subsets/Paperlogy-dynamic-subset.css" rel="stylesheet">
<style>body{margin:0;background:#FAF9F7;}</style></head><body><div id="root"></div></body></html>`;
}

function paintPreview() {
  const f = document.getElementById('preview');
  const doc = f.contentDocument; if (!doc) return;
  const win = f.contentWindow;
  const y = win ? win.scrollY : 0, x = win ? win.scrollX : 0;   // 현재 스크롤 보존
  const root = doc.getElementById('root') || doc.body;
  root.innerHTML = renderDetailHTML(spec, Catalog);
  if (win) win.scrollTo(x, y);                                   // 그 자리에 복원
}

function updatePreview() {
  clearTimeout(previewTimer);
  previewTimer = setTimeout(() => { if (previewReady) paintPreview(); }, 120);
}

/* ════════════ 상단바 (토글 / 내보내기 / 불러오기) ════════════ */
function bindTopbar() {
  document.getElementById('viewPC').addEventListener('click', () => setView('pc'));
  document.getElementById('viewMobile').addEventListener('click', () => setView('mobile'));
  document.getElementById('btnHTML').addEventListener('click', copyHTML);
  document.getElementById('btnSpec').addEventListener('click', downloadSpec);
  document.getElementById('fileSpec').addEventListener('change', importSpec);
}
function setView(mode) {
  const frame = document.getElementById('previewFrame');
  frame.classList.toggle('mobile', mode === 'mobile');
  document.getElementById('viewPC').classList.toggle('on', mode === 'pc');
  document.getElementById('viewMobile').classList.toggle('on', mode === 'mobile');
}
function copyHTML() {
  const html = renderDetailHTML(spec, Catalog);
  navigator.clipboard.writeText(html).then(
    () => toast('완성 HTML 복사됨 — 아임웹 상품 상세에 붙여넣기'),
    () => { fallbackCopy(html); toast('복사됨 (fallback)'); }
  );
}
function fallbackCopy(text) {
  const ta = document.createElement('textarea'); ta.value = text; document.body.appendChild(ta);
  ta.select(); document.execCommand('copy'); ta.remove();
}
function downloadSpec() {
  spec.updated_at = todayStr();
  if (!spec.created_at) spec.created_at = spec.updated_at;
  const blob = new Blob([JSON.stringify(spec, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `${spec.model || 'model'}.json`;
  a.click(); URL.revokeObjectURL(a.href);
  toast(`${a.download} 다운로드 — projects/products/specs/ 에 보관`);
}
function importSpec(e) {
  const file = e.target.files[0]; if (!file) return;
  const r = new FileReader();
  r.onload = () => {
    try {
      const loaded = JSON.parse(r.result);
      spec = Object.assign(defaultSpec(loaded.type || 'blunt'), loaded);
      spec.selections = spec.selections || {}; spec.copy_selections = spec.copy_selections || {};
      spec.custom_fields = spec.custom_fields || {}; spec.spec_meta = spec.spec_meta || {};
      renderPanel(); updatePreview(); toast(`${spec.model || 'spec'} 불러옴`);
    } catch (err) { toast('불러오기 실패: ' + err.message); }
  };
  r.readAsText(file); e.target.value = '';
}
function todayStr() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}
let toastTimer = null;
function toast(msg) {
  const t = document.getElementById('toast'); t.textContent = msg; t.classList.add('on');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('on'), 3200);
}
