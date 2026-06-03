/* ──────────────────────────────────────────────────────────────
   workspace.js — 카탈로그 디자인 작업대
   빌더와 동일한 catalog.js + template.js 렌더러로 모든 카드를 실제 형태로 출력.
   ★ PC 패널 + 모바일 패널(iframe 390px)을 한 화면에 동시 표시 (토글 없음).
     각 iframe 안의 vw 가 그 패널 폭을 읽음 → PC/모바일 둘 다 실제와 일치.
     두 패널 스크롤 동기화로 같은 카드를 나란히 비교.
   여기서 정하면 빌더에 자동 반영 (같은 catalog+template).
   ────────────────────────────────────────────────────────────── */

const WS_CSS = `
  :root{ --void:#1A1A1A; --mute:#8A8580; --line:#EDEBE8; --shell:#F5F3F0; --cream:#FAF9F7; }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{ font-family:'Plus Jakarta Sans','Noto Sans KR',sans-serif; background:#E8E6E2; color:var(--void); -webkit-font-smoothing:antialiased; }
  .ws-wrap{ margin:0 auto; padding:20px 16px 140px; }
  .ws-h2{ font-family:'Outfit',sans-serif; font-size:12px; font-weight:800; letter-spacing:.16em; text-transform:uppercase; color:var(--void); margin:28px 4px 12px; padding-bottom:7px; border-bottom:2px solid var(--void); }
  .ws-h2:first-child{ margin-top:4px; }
  .ws-note{ font-size:11.5px; color:#6b6862; line-height:1.7; background:#fff; border:1px solid var(--line); border-radius:10px; padding:12px 14px; margin-bottom:4px; }
  .ws-note b{ color:var(--void); }
  .ws-card{ background:#fff; border:1px solid var(--line); border-radius:12px; overflow:hidden; margin-bottom:16px; }
  .ws-card__bar{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; padding:10px 14px; border-bottom:1px solid var(--line); background:#fbfaf9; }
  .ws-card__id{ font-family:'Outfit',sans-serif; font-size:11.5px; font-weight:800; color:var(--void); }
  .ws-card__label{ font-size:11.5px; color:var(--mute); }
  .ws-card__applies{ margin-left:auto; font-size:9.5px; font-weight:700; letter-spacing:.05em; color:#fff; background:#B8B4AF; border-radius:999px; padding:3px 9px; }
  .ws-stage{ background:var(--cream); padding:clamp(16px,3vw,26px); }
  .ws-copy{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:12px 14px; margin-bottom:12px; }
  .ws-copy__bar{ display:flex; align-items:center; gap:8px; margin-bottom:8px; }
  .ws-copy__opt{ font-size:12.5px; color:#2D2D2D; line-height:1.6; padding:6px 0; border-top:1px dashed var(--line); white-space:pre-line; }
  .ws-copy__opt:first-of-type{ border-top:0; }
  .ws-copy__id{ font-family:'Outfit',sans-serif; font-size:9.5px; color:#a8a49e; margin-right:2px; }
`;

const FONT_LINKS =
  '<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@700;800;900&family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Noto+Sans+KR:wght@400;500;700&display=swap" rel="stylesheet">'
  + '<link href="https://cdn.jsdelivr.net/gh/fonts-archive/Paperlogy/subsets/Paperlogy-dynamic-subset.css" rel="stylesheet">';

window.addEventListener('DOMContentLoaded', async () => {
  try {
    await Catalog.load();
  } catch (e) {
    paint('<div style="padding:40px 20px;color:#8A8580;font-size:13px;">카탈로그 로드 실패: ' + e.message + '<br><small>정적 서버로 열어야 합니다 (preview.bat).</small></div>');
    return;
  }
  render();
});

/* 두 iframe(PC·모바일)에 동일 콘텐츠 주입 + 로드 후 스크롤 동기화 */
function paint(innerHtml) {
  const doc = '<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + FONT_LINKS + '<style>' + WS_CSS + '</style></head><body><div class="ws-wrap">' + innerHtml + '</div></body></html>';
  const pc = document.getElementById('wsPC_frame');
  const mo = document.getElementById('wsMo_frame');
  let pending = 2;
  const onload = () => { if (--pending === 0) wireScrollSync(pc, mo); };
  pc.onload = onload; mo.onload = onload;
  pc.srcdoc = doc; mo.srcdoc = doc;
}

function wireScrollSync(pc, mo) {
  let lock = false;
  const link = (src, dst) => {
    const win = src.contentWindow;
    if (!win) return;
    win.addEventListener('scroll', () => {
      if (lock) return; lock = true;
      try {
        const sDoc = src.contentDocument.documentElement;
        const dDoc = dst.contentDocument.documentElement;
        const sMax = sDoc.scrollHeight - src.clientHeight;
        const dMax = dDoc.scrollHeight - dst.clientHeight;
        const r = sMax > 0 ? win.scrollY / sMax : 0;
        dst.contentWindow.scrollTo(0, Math.max(0, dMax) * r);
      } catch (e) {}
      requestAnimationFrame(() => { lock = false; });
    }, { passive: true });
  };
  link(pc, mo); link(mo, pc);
}

function firstId(cardType) {
  const c = Catalog.byCardType[cardType];
  return c && c.options && c.options[0] ? c.options[0].id : '';
}
function variantOf(cardType) {
  if (cardType === 'blade_edge') return 'edge';
  if (cardType === 'blade_design') return 'design';
  return 'text';
}
function esc2(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }

/* 패널 = 헤더바(id · 라벨 · 적용종류) + 무대(실제 렌더) — 군더더기 제거 */
function panel(idText, labelText, appliesArr, innerHtml) {
  return '<div class="ws-card">'
    + '<div class="ws-card__bar">'
    + '<span class="ws-card__id">' + esc2(idText) + '</span>'
    + '<span class="ws-card__label">' + esc2(labelText) + '</span>'
    + '<span class="ws-card__applies">' + (appliesArr || []).join(' · ') + '</span>'
    + '</div>'
    + '<div class="ws-stage">' + innerHtml + '</div>'
    + '</div>';
}

function cardPanel(cardType) {
  const card = Catalog.byCardType[cardType];
  if (!card) return '';
  const label = (card.label_ko || '') + (card.label_subtitle_ko ? ' · ' + card.label_subtitle_ko : '');
  return panel(cardType, label, card.applies_to, cardGroup(card, firstId(cardType), variantOf(cardType)));
}

function copyPanel(pool) {
  const opts = (pool.options || []).map(o =>
    '<div class="ws-copy__opt"><span class="ws-copy__id">' + esc2(o.id) + '</span>' + esc2(o.text || '') + '</div>'
  ).join('');
  return '<div class="ws-copy">'
    + '<div class="ws-copy__bar">'
    + '<span class="ws-card__id">' + esc2(pool.copy_type) + '</span>'
    + '<span class="ws-card__label">' + esc2(pool.label_ko || '') + '</span>'
    + '<span class="ws-card__applies">' + (pool.applies_to || []).join(' · ') + '</span>'
    + '</div>' + opts
    + '</div>';
}

function render() {
  let h = '';
  h += '<div class="ws-note">빌더와 <b>같은 catalog + template</b> 실제 렌더. 좌=PC / 우=모바일(390px) 동시.<br>'
    + '아이콘·이름·설명 → <b>catalog/cards</b> · 섹션 문구 → <b>catalog/copy_pool</b> · 카드 형태 → <b>template.js</b></div>';

  const sampleSpec = {
    type: 'blunt', price_grade: 'A',
    selections: { handle_grip: firstId('handle_grip'), handle_camel: firstId('handle_camel') },
  };

  h += '<div class="ws-h2">① 공통 카드 (전 종류 공용)</div>';
  if (Catalog.byCardType['handle_grip'] || Catalog.byCardType['handle_camel']) {
    h += panel('handle', '핸들 (Grip + Camel)', ['blunt', 'thinning', 'long', 'dry'], handleGroup(sampleSpec, Catalog));
  }
  if (Catalog.byCardType['grade']) {
    const gradeInner = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:clamp(10px,1.5vw,16px)">' + gradeCards(sampleSpec, Catalog) + '</div>';
    h += panel('grade', Catalog.byCardType['grade'].label_ko || 'GRADE', Catalog.byCardType['grade'].applies_to, gradeInner);
  }

  h += '<div class="ws-h2">② 가위 종류별 특징 카드</div>';
  ['blade_edge', 'blade_design', 'thinning_teeth', 'thinning_holes', 'thinning_reduction', 'dry_cutting_style']
    .forEach(ct => { h += cardPanel(ct); });

  h += '<div class="ws-h2">③ 섹션 문구 풀</div>';
  Catalog.copy.forEach(pool => { h += copyPanel(pool); });

  paint(h);
}
