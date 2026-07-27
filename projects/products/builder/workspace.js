/* ──────────────────────────────────────────────────────────────
   workspace.js — 카탈로그 디자인 작업대 (PC + 모바일 동시 미리보기)
   · 빌더와 동일한 catalog.js + template.js 렌더러로 모든 카드를 실제 형태로 출력.
   · iframe(좌 PC, 우 모바일 390px) → vw 가 각 패널 폭을 읽어 정확.
   · ★ 편집기(editor.html)와 BroadcastChannel 로 연결 → 편집기 타이핑이 여기 실시간 반영
     (저장·서버 불필요, in-place 갱신이라 깜빡임/스크롤 튐 없음).
   ────────────────────────────────────────────────────────────── */

const WS_CSS = `
  :root{ --void:#1A1A1A; --mute:#8A8580; --line:#EDEBE8; --shell:#F5F3F0; --cream:#FAF9F7; }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{ font-family:'Plus Jakarta Sans','Noto Sans KR',sans-serif; background:#E8E6E2; color:var(--void); -webkit-font-smoothing:antialiased; }
  .ws-wrap{ max-width:840px; margin:0 auto; padding:20px 16px 140px; } /* 실제 상품 상세 본문 폭(840px) */
  .ws-h2{ font-family:'Outfit',sans-serif; font-size:12px; font-weight:800; letter-spacing:.16em; text-transform:uppercase; color:var(--void); margin:28px 4px 12px; padding-bottom:7px; border-bottom:2px solid var(--void); }
  .ws-h2:first-child{ margin-top:4px; }
  .ws-note{ font-size:11.5px; color:#6b6862; line-height:1.7; background:#fff; border:1px solid var(--line); border-radius:10px; padding:12px 14px; margin-bottom:4px; }
  .ws-note b{ color:var(--void); }
  .ws-card{ background:#fff; border:1px solid var(--line); border-radius:12px; overflow:hidden; margin-bottom:16px; }
  .ws-card__bar{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; padding:10px 14px; border-bottom:1px solid var(--line); background:#fbfaf9; }
  .ws-card__id{ font-family:'Outfit',sans-serif; font-size:11.5px; font-weight:800; color:var(--void); }
  .ws-card__label{ font-size:11.5px; color:var(--mute); }
  .ws-card__applies{ margin-left:auto; font-family:'Outfit',sans-serif; font-size:15px; font-weight:900; letter-spacing:.08em; text-transform:uppercase; color:var(--cream); background:var(--void); border-radius:8px; padding:6px 16px; }
  /* for_you 계열: 옵션 각각 카드로 그리면 혼잡 → 한 그룹에 카테고리별 한 줄씩 */
  .ws-fy{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:6px 16px 14px; }
  .ws-fy__cat{ font-family:'Outfit',sans-serif; font-size:10.5px; font-weight:800; letter-spacing:.14em; text-transform:uppercase; color:var(--mute); margin:16px 0 6px; padding-bottom:5px; border-bottom:1px solid var(--line); }
  .ws-fy__row{ display:flex; align-items:baseline; gap:10px; padding:6px 0; }
  .ws-fy__id{ flex:0 0 auto; font-family:'Outfit',sans-serif; font-size:9px; font-weight:700; color:#c4c0ba; letter-spacing:.03em; min-width:76px; }
  .ws-fy__txt{ flex:1 1 auto; font-size:14px; color:#2D2D2D; line-height:1.5; }
  .ws-fy__txt.empty{ color:#c4c0ba; font-style:italic; }
  .ws-stage{ background:var(--cream); padding:clamp(16px,3vw,26px); }
  .ws-copy{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:12px 14px; margin-bottom:12px; }
  .ws-copy__bar{ display:flex; align-items:center; gap:8px; margin-bottom:8px; }
  /* 문구 = 실제 페이지 모양 그대로 렌더 (id 는 좌측 옅은 라벨로만) */
  .ws-copy__stage{ display:flex; flex-direction:column; gap:14px; }
  .ws-copy__row{ padding:10px 0; border-top:1px dashed var(--line); }
  .ws-copy__row:first-child{ border-top:0; padding-top:0; }
  .ws-copy__id{ font-family:'Outfit',sans-serif; font-size:9.5px; font-weight:700; color:#b8b4af; letter-spacing:.04em; margin-bottom:6px; }
  .ws-copy__empty{ font-size:12px; color:#b8b4af; font-style:italic; }
  .ws-card[data-src],.ws-copy[data-src]{ cursor:pointer; }
  .ws-card[data-src]:hover,.ws-copy[data-src]:hover{ border-color:var(--void); box-shadow:0 0 0 3px rgba(26,26,26,.06); }
`;

const FONT_LINKS =
  '<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Noto+Sans+KR:wght@400;500;700&display=swap" rel="stylesheet">'
  + '<link href="https://cdn.jsdelivr.net/gh/fonts-archive/Paperlogy/subsets/Paperlogy-dynamic-subset.css" rel="stylesheet">'
  + '<link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard-dynamic-subset.css" rel="stylesheet">';

const SKELETON = '<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
  + FONT_LINKS + '<style>' + WS_CSS + '</style></head><body><div id="root" class="ws-wrap"></div></body></html>';

let pcReady = false, moReady = false, repaintTimer = null;

window.addEventListener('DOMContentLoaded', async () => {
  try {
    await Catalog.load();
  } catch (e) {
    const f = document.getElementById('wsPC_frame');
    if (f) f.srcdoc = '<div style="padding:40px;font-family:sans-serif;color:#8A8580">카탈로그 로드 실패: ' + e.message + ' (preview.bat 로 열어야 함)</div>';
    return;
  }
  NEUTRAL = true;       // 작업대 기본 = 정보 보기(모든 옵션 또렷)
  bindModeToggle();
  initFrames();
  listenChannel();
});

/* 두 iframe 을 스켈레톤으로 1회 로드 → 이후 #root.innerHTML 만 교체(깜빡임 없음) */
function initFrames() {
  const pc = document.getElementById('wsPC_frame');
  const mo = document.getElementById('wsMo_frame');
  pc.addEventListener('load', () => { pcReady = true; attachClick(pc); paintInto(pc); maybeSync(); }, { once: true });
  mo.addEventListener('load', () => { moReady = true; attachClick(mo); paintInto(mo); maybeSync(); }, { once: true });
  pc.srcdoc = SKELETON;
  mo.srcdoc = SKELETON;
}

/* 미리보기 패널 클릭 → 최상위(스튜디오)로 신호 → 왼쪽 편집기가 해당 항목으로 스크롤.
   스튜디오 iframe 안일 때만 동작(window.top !== window). 단독 작업대에선 무해. */
function attachClick(frame) {
  try {
    const doc = frame.contentDocument;
    if (!doc || window.top === window) return;
    doc.addEventListener('click', (e) => {
      const el = e.target.closest && e.target.closest('[data-src]');
      if (!el) return;
      const src = el.getAttribute('data-src');
      if (src) { try { window.top.postMessage({ type: 'focus-src', src }, '*'); } catch (_) {} }
    });
  } catch (e) {}
}
function maybeSync() { if (pcReady && moReady) wireScrollSync(document.getElementById('wsPC_frame'), document.getElementById('wsMo_frame')); }

/* 정보 보기(중립) ↔ 선택 미리보기 토글 */
function bindModeToggle() {
  const n = document.getElementById('wsNeutral');
  const s = document.getElementById('wsSelected');
  if (!n || !s) return;
  n.onclick = () => { NEUTRAL = true; n.classList.add('on'); s.classList.remove('on'); paintAll(); };
  s.onclick = () => { NEUTRAL = false; s.classList.add('on'); n.classList.remove('on'); paintAll(); };
}

function paintInto(frame) {
  try {
    const doc = frame.contentDocument;
    if (!doc) return;
    const root = doc.getElementById('root');
    if (!root) return;
    const y = frame.contentWindow.scrollY;
    root.innerHTML = buildContent();
    frame.contentWindow.scrollTo(0, y); // 스크롤 유지
  } catch (e) {}
}
function paintAll() { paintInto(document.getElementById('wsPC_frame')); paintInto(document.getElementById('wsMo_frame')); }
function schedulePaint() { clearTimeout(repaintTimer); repaintTimer = setTimeout(paintAll, 100); }

function wireScrollSync(pc, mo) {
  let lock = false;
  const link = (src, dst) => {
    const win = src.contentWindow;
    if (!win) return;
    win.addEventListener('scroll', () => {
      if (lock) return; lock = true;
      try {
        const sMax = src.contentDocument.documentElement.scrollHeight - src.clientHeight;
        const dMax = dst.contentDocument.documentElement.scrollHeight - dst.clientHeight;
        const r = sMax > 0 ? win.scrollY / sMax : 0;
        dst.contentWindow.scrollTo(0, Math.max(0, dMax) * r);
      } catch (e) {}
      requestAnimationFrame(() => { lock = false; });
    }, { passive: true });
  };
  link(pc, mo); link(mo, pc);
}

/* 편집기 실시간 연동 — 편집기에서 {src,path,value} 오면 Catalog 갱신 후 재페인트 */
function listenChannel() {
  try {
    const ch = new BroadcastChannel('mamoru-catalog');
    ch.onmessage = (e) => {
      const d = e.data || {};
      if (!d.src || !d.path) return;
      const obj = Catalog.cards.concat(Catalog.copy).find(o => o._src === d.src);
      if (obj) { setByPath(obj, d.path, d.value); schedulePaint(); }
    };
  } catch (e) {}
}
function setByPath(obj, pathStr, val) {
  const keys = pathStr.split('.');
  let cur = obj;
  for (let i = 0; i < keys.length - 1; i++) cur = cur[keys[i]];
  cur[keys[keys.length - 1]] = val;
}

/* ── 콘텐츠 빌드 (① 공통 ② 종류별 ③ 문구풀) ── */
function firstId(cardType) {
  const c = Catalog.byCardType[cardType];
  return c && c.options && c.options[0] ? c.options[0].id : '';
}
function variantOf(cardType) {
  if (cardType === 'blade_edge' || cardType === 'blade_edge_long' || cardType === 'blade_edge_dry') return 'edge';
  if (cardType === 'blade_design') return 'design';
  return 'text';
}
function esc2(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }

function panel(idText, labelText, appliesArr, innerHtml, src) {
  return '<div class="ws-card"' + (src ? ' data-src="' + esc2(src) + '"' : '') + '><div class="ws-card__bar">'
    + '<span class="ws-card__id">' + esc2(idText) + '</span>'
    + '<span class="ws-card__label">' + esc2(labelText) + '</span>'
    + '<span class="ws-card__applies">' + (appliesArr || []).join(' · ') + '</span>'
    + '</div><div class="ws-stage">' + innerHtml + '</div></div>';
}
function cardPanel(cardType) {
  const card = Catalog.byCardType[cardType];
  if (!card) return '';
  const label = (card.label_ko || '') + (card.label_subtitle_ko ? ' · ' + card.label_subtitle_ko : '');
  return panel(cardType, label, card.applies_to, cardGroup(card, firstId(cardType), variantOf(cardType)), card._src);
}
/* 문구 풀 = 실제 상세페이지에 나갈 '그 모양 그대로' 미리보기.
   template.js 의 copyBlock() 을 그대로 쓴다(출력 SSOT) → 여기서 본 모양 = 고객이 볼 모양. */
/* for_you_match / for_you_miss = 옵션이 많아 각각 forYouCard 로 그리면 "이런 분에게 맞습니다"가
   반복되며 혼잡 → 한 그룹에 카테고리별 한 줄씩(길면 자동 2줄) 컴팩트 리스트로. */
function forYouListPreview(pool) {
  const opts = pool.options || [];
  const order = [], byCat = {};
  opts.forEach(o => {
    const cat = o.category || '기타';
    if (!byCat[cat]) { byCat[cat] = []; order.push(cat); }
    byCat[cat].push(o);
  });
  let rows = '';
  order.forEach(cat => {
    rows += '<div class="ws-fy__cat">' + esc2(cat) + '</div>';
    byCat[cat].forEach(o => {
      const t = (o.text || '').trim();
      rows += '<div class="ws-fy__row"><span class="ws-fy__id">' + esc2(o.id) + '</span>'
        + '<span class="ws-fy__txt' + (t ? '' : ' empty') + '">' + (t ? esc2(t) : '(빈 문구)') + '</span></div>';
    });
  });
  return '<div class="ws-fy">' + (rows || '<div class="ws-copy__empty">문구 없음</div>') + '</div>';
}
function copyPanel(pool) {
  if (pool.copy_type === 'for_you_match' || pool.copy_type === 'for_you_miss') {
    return panel(pool.copy_type, pool.label_ko || '', pool.applies_to, forYouListPreview(pool), pool._src);
  }
  const opts = (pool.options || []).map(o => {
    const text = (o.text || '').trim();
    const rendered = text
      ? copyBlock(pool.copy_type, o.text)
      : '<div class="ws-copy__empty">(빈 문구 — 편집기에서 채우면 여기 바로 보입니다)</div>';
    return '<div class="ws-copy__row">'
      + '<div class="ws-copy__id">' + esc2(o.id) + '</div>'
      + '<div class="ws-copy__render">' + rendered + '</div>'
      + '</div>';
  }).join('');
  const inner = '<div class="ws-copy__stage">' + (opts || '<div class="ws-copy__empty">문구 없음</div>') + '</div>';
  return panel(pool.copy_type, pool.label_ko || '', pool.applies_to, inner, pool._src);
}

const COMMON_CARDS = ['handle_grip', 'handle_camel', 'grade'];
const TYPE_CARDS = ['blade_edge', 'blade_edge_long', 'blade_edge_dry', 'blade_design', 'thinning_teeth', 'thinning_holes', 'thinning_reduction'];

function buildContent() {
  let h = '<div class="ws-note">편집기에서 문구 수정 시 <b>여기 실시간 반영</b> · 좌=PC / 우=모바일(390px) · 스크롤 동기화</div>';
  const sampleSpec = { type: 'blunt', price_grade: 'A', selections: { handle_grip: firstId('handle_grip'), handle_camel: firstId('handle_camel') } };

  h += '<div class="ws-h2">① 공통 카드</div>';
  if (Catalog.byCardType['handle_grip'] || Catalog.byCardType['handle_camel']) {
    h += panel('handle', '핸들 (Grip + Camel)', ['blunt', 'thinning', 'long', 'dry'], handleGroup(sampleSpec, Catalog), (Catalog.byCardType['handle_grip'] || {})._src);
  }
  if (Catalog.byCardType['grade']) {
    const gradeInner = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:clamp(10px,1.5vw,16px)">' + gradeCards(sampleSpec, Catalog) + '</div>';
    h += panel('grade', Catalog.byCardType['grade'].label_ko || 'GRADE', Catalog.byCardType['grade'].applies_to, gradeInner, Catalog.byCardType['grade']._src);
  }

  h += '<div class="ws-h2">② 가위 종류별 특징 카드</div>';
  TYPE_CARDS.forEach(ct => { h += cardPanel(ct); });

  h += '<div class="ws-h2">③ 섹션 문구 풀</div>';
  Catalog.copy.forEach(pool => { h += copyPanel(pool); });
  return h;
}
