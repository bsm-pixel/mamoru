/* ──────────────────────────────────────────────────────────────
   workspace.js — 카탈로그 디자인 작업대
   빌더와 동일한 catalog.js(로더) + template.js(렌더러)를 그대로 사용.
   모든 카드 종류 × 옵션을 실제 페이지와 같은 형태로 한 화면에 출력 →
   카드 표시 형태 / 아이콘(SVG) / 문구를 보면서 일괄 디자인.
   여기서 정하면 빌더에 자동 반영 (같은 catalog+template).
   ────────────────────────────────────────────────────────────── */

window.addEventListener('DOMContentLoaded', async () => {
  try {
    await Catalog.load();
  } catch (e) {
    document.getElementById('ws').innerHTML = '<div class="ws-loading">카탈로그 로드 실패: ' + e.message + '<br><small>정적 서버로 열어야 합니다 (file:// 차단)</small></div>';
    return;
  }
  bindToggle();
  render();
});

function bindToggle() {
  const wrap = document.getElementById('ws');
  const pc = document.getElementById('wsPC');
  const mo = document.getElementById('wsMobile');
  pc.onclick = () => { wrap.classList.remove('mobile'); pc.classList.add('on'); mo.classList.remove('on'); };
  mo.onclick = () => { wrap.classList.add('mobile'); mo.classList.add('on'); pc.classList.remove('on'); };
}

function firstId(cardType) {
  const c = Catalog.byCardType[cardType];
  return c && c.options && c.options[0] ? c.options[0].id : '';
}

/* template.js 의 카드별 변형 매핑 (renderDetailHTML 과 동일 규칙) */
function variantOf(cardType) {
  if (cardType === 'blade_edge') return 'edge';
  if (cardType === 'blade_design') return 'design';
  return 'text';
}

function esc2(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }

function panel(idText, labelText, appliesArr, fileText, innerHtml) {
  return '<div class="ws-card">'
    + '<div class="ws-card__bar">'
    + '<span class="ws-card__id">' + esc2(idText) + '</span>'
    + '<span class="ws-card__label">' + esc2(labelText) + '</span>'
    + '<span class="ws-card__applies">' + (appliesArr || []).join(' · ') + '</span>'
    + '</div>'
    + '<div class="ws-card__bar" style="border-top:0;background:#fff"><span class="ws-card__file">' + esc2(fileText) + '</span></div>'
    + '<div class="ws-stage">' + innerHtml + '</div>'
    + '</div>';
}

function cardPanel(cardType) {
  const card = Catalog.byCardType[cardType];
  if (!card) return '';
  const label = (card.label_ko || '') + (card.label_subtitle_ko ? ' · ' + card.label_subtitle_ko : '');
  const file = 'catalog/' + (card._src || ('cards/' + cardType + '.json'));
  return panel(cardType, label, card.applies_to, file, cardGroup(card, firstId(cardType), variantOf(cardType)));
}

function copyPanel(pool) {
  const opts = (pool.options || []).map(o =>
    '<div class="ws-copy__opt"><span class="ws-copy__id">' + esc2(o.id) + '</span>  ' + esc2(o.text || '') + '</div>'
  ).join('');
  return '<div class="ws-copy">'
    + '<div class="ws-copy__bar">'
    + '<span class="ws-card__id">' + esc2(pool.copy_type) + '</span>'
    + '<span class="ws-card__label">' + esc2(pool.label_ko || '') + '</span>'
    + '<span class="ws-card__applies">' + (pool.applies_to || []).join(' · ') + '</span>'
    + '</div>' + opts
    + '<div class="ws-card__file" style="margin-top:10px">catalog/' + esc2(pool._src || '') + '</div>'
    + '</div>';
}

function render() {
  const ws = document.getElementById('ws');
  let h = '';

  h += '<div class="ws-note">이 작업대는 빌더와 <b>같은 catalog + template</b>으로 모든 카드를 실제 렌더합니다. 여기서 정하면 빌더에 자동 반영.<br>'
    + '· <b>아이콘(SVG)·이름·설명</b> → catalog/cards/*.json &nbsp;&nbsp;'
    + '· <b>섹션 문구</b> → catalog/copy_pool/*.json &nbsp;&nbsp;'
    + '· <b>카드 표시 형태</b> → builder/template.js</div>';

  const sampleSpec = {
    type: 'blunt', price_grade: 'A',
    selections: { handle_grip: firstId('handle_grip'), handle_camel: firstId('handle_camel') },
  };

  // ① 공통 카드
  h += '<div class="ws-h2">① 공통 카드 (전 종류 공용)</div>';
  if (Catalog.byCardType['handle_grip'] || Catalog.byCardType['handle_camel']) {
    h += panel('handle', '핸들 — 손에 맞추다 (Grip + Camel)', ['blunt', 'thinning', 'long', 'dry'],
      'catalog/cards/handle_grip.json + handle_camel.json', handleGroup(sampleSpec, Catalog));
  }
  if (Catalog.byCardType['grade']) {
    const gradeInner = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:clamp(12px,1.5vw,20px)">' + gradeCards(sampleSpec, Catalog) + '</div>';
    h += panel('grade', (Catalog.byCardType['grade'].label_ko || 'GRADE'), Catalog.byCardType['grade'].applies_to,
      'catalog/' + (Catalog.byCardType['grade']._src || 'cards/grade.json'), gradeInner);
  }

  // ② 종류별 특징 카드
  h += '<div class="ws-h2">② 가위 종류별 특징 카드</div>';
  ['blade_edge', 'blade_design', 'thinning_teeth', 'thinning_holes', 'thinning_reduction', 'dry_cutting_style']
    .forEach(ct => { h += cardPanel(ct); });

  // ③ 섹션 문구 풀
  h += '<div class="ws-h2">③ 섹션 문구 풀</div>';
  Catalog.copy.forEach(pool => { h += copyPanel(pool); });

  ws.innerHTML = h;
}
