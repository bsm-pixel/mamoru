/* ──────────────────────────────────────────────────────────────
   template.js — 출력 HTML 생성기 (SSOT)
   renderDetailHTML(spec, catalog) → v10_trendy 레이아웃 inline-style HTML 문자열
   · 미리보기 iframe 과 "HTML 복사" 가 모두 이 결과를 사용 → WYSIWYG 보장
   · v10_trendy.html = 디자인 레퍼런스. 실제 생성은 이 파일이 담당.
   · 아임웹 상품 상세 = inline style 전용 (style/script/iframe 금지) 룰 준수
   ────────────────────────────────────────────────────────────── */

const IMG_HOST = 'https://page.mamoru.kr';
const TYPE_LABEL = { blunt: 'Blunt', thinning: 'Thinning', long: 'Long', dry: 'Dry' };
// 작업대 '정보 보기(중립)' 모드: true 면 모든 옵션을 또렷하게(선택 강조/흐림 없이) 렌더. 실제 페이지/빌더는 false.
var NEUTRAL = false;

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function nl2br(s) { return esc(s).replace(/\n/g, '<br>'); }

/* 서체 규칙 (사장님 확정 2026-07-12): 영문·숫자 = Paperlogy / 한글 = Pretendard.
   ⚠️ Paperlogy에는 한글도 들어 있어서 스택 앞에 두면 한글까지 Paperlogy로 나온다.
      → 글자 종류별로 span을 갈라 각각 지정해야 한다(스택만으로는 분리 불가). */
const FONT_EN = "'Paperlogy','Outfit',sans-serif";
const FONT_KO = "'Pretendard','Noto Sans KR',sans-serif";
const HANGUL = /[가-힣ㄱ-ㅎㅏ-ㅣ]/;

/* 큰 표시값에서 뒤 단위(발/홈/% 등)를 자동으로 2/3 크기로 — "26발" → 26 크게 + 발 작게.
   숫자는 Paperlogy, 한글 단위는 Pretendard로 분리 렌더 */
function bigValue(str) {
  const s = String(str == null ? '' : str);
  const m = s.match(/^([\d.]+)(.+)$/);   // 앞 숫자 + 뒤 단위
  if (!m) {
    // 숫자로 시작하지 않는 값(F·A·G 같은 영문 코드, 순한글 등)
    return HANGUL.test(s) ? `<span style="font-family:${FONT_KO};">${esc(s)}</span>` : esc(s);
  }
  const unitFont = HANGUL.test(m[2]) ? FONT_KO : FONT_EN;   // '홈'·'발'=한글 / '%'·'mm'=영문
  return `${esc(m[1])}<span style="font-size:0.62em;font-family:${unitFont};">${esc(m[2])}</span>`;
}

/* 카피 풀 placeholder({길이}/{날선}/{날등}) 치환 */
function fillPlaceholders(text, spec, catalog) {
  if (!text) return '';
  const edge = catalog.cardOption('blade_edge', spec.selections?.blade_edge);
  const design = catalog.cardOption('blade_design', spec.selections?.blade_design);
  return text
    .replace(/\{길이\}/g, spec.size_inch != null ? spec.size_inch : '')
    .replace(/\{날선\}/g, edge ? edge.name_ko : '')
    .replace(/\{날등\}/g, design ? design.name_ko : '');
}

/* spec + catalog 에서 카피 텍스트 해석 (custom 우선, 없으면 풀 id → text) */
function resolveCopy(spec, catalog, copyType, opts = {}) {
  const cf = spec.custom_fields || {};
  const cs = spec.copy_selections || {};
  // custom 직접입력 우선
  const customKey = opts.customKey;
  if (customKey && cf[customKey]) return fillPlaceholders(cf[customKey], spec, catalog);
  const id = cs[copyType];
  if (id == null) return '';
  if (Array.isArray(id)) {
    return id.map(i => fillPlaceholders(catalog.copyText(copyType, spec.type, i), spec, catalog));
  }
  return fillPlaceholders(catalog.copyText(copyType, spec.type, id), spec, catalog);
}

function imgURL(spec, file) {
  if (!file) return '';
  const host = spec.image_host || IMG_HOST;
  const folder = spec.images_folder || spec.model || 'MODEL';
  return `${host}/projects/products/product_detail/${folder}/images/${file}`;
}

/* ─── 섹션 라벨 (— NN / TITLE) ─── */
function eyebrow(num, title, dark) {
  const col = dark ? 'rgba(245,245,243,0.5)' : '#8A8580';
  return `<div id="sec-${num}" style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:700;color:${col};letter-spacing:0.25em;margin-bottom:clamp(24px,3vw,40px);scroll-margin-top:20px;">— ${num} / ${esc(title)}</div>`;
}

/* ─── 카드 1장 (BLADE EDGE/DESIGN/text 변형) ───
   variant: 'edge'(svg 100%) | 'design'(svg 65% center) | 'text'(svg 없음)
   card: 카드 정의(number_only 판별용 — 없으면 기존 동작) */
function optionCard(opt, selected, variant, card) {
  const dark = selected && !NEUTRAL; // 다크 강조는 선택+비중립일 때만
  const bg = dark ? 'background:#1A1A1A;color:#FAF9F7;' : 'background:#FFFFFF;border:1px solid #EDEBE8;';
  const letterCol = dark ? '#FAF9F7' : (NEUTRAL ? '#1A1A1A' : '#8A8580');
  const nameCol = dark ? '#FAF9F7' : (NEUTRAL ? '#1A1A1A' : '#8A8580');
  const enCol = dark ? 'rgba(245,245,243,0.45)' : (NEUTRAL ? '#8A8580' : '#B8B4AF');
  const divCol = dark ? 'rgba(245,245,243,0.15)' : '#EDEBE8';
  const descCol = dark ? 'rgba(245,245,243,0.75)' : (NEUTRAL ? '#2D2D2D' : '#B8B4AF');
  const check = dark ? `<span style="font-size:clamp(14px,2vw,20px);color:#FAF9F7;font-weight:700;line-height:1;flex-shrink:0;">✓</span>` : '';

  /* number_only 카드(빗살 수·홈 수) = 숫자만 카드 정중앙.
     단위(발·홈)는 카드 위 라벨이 이미 말해주므로 카드마다 반복하면 중복 + 시선만 분산된다.
     구분선·설명도 없앤다(내용이 없어 빈 줄만 남던 자리). */
  if (card && card.number_only) {
    const num = opt.value != null ? opt.value : (String(opt.id).match(/[\d.]+/) || [opt.id])[0];
    return `<div style="${bg}border-radius:clamp(8px,1.2vw,12px);padding:clamp(20px,3vw,34px) clamp(8px,1.5vw,16px);text-align:center;">
      <span style="font-family:${FONT_EN};font-size:clamp(30px,6vw,68px);font-weight:900;color:${letterCol};line-height:1;letter-spacing:-0.02em;">${esc(String(num))}</span>
    </div>`;
  }

  let svg = '';
  if (opt.svg_inline) {
    const op = (dark || NEUTRAL) ? '' : 'opacity:0.35;';
    // design=중앙 65% / edge=100% / text(틴닝 등)=아이콘 크기(홈형태 등 표현)
    const w = variant === 'design' ? 'width:65%;margin:0 auto clamp(16px,2.5vw,28px);'
            : variant === 'edge' ? 'width:100%;margin-bottom:clamp(16px,2.5vw,28px);'
            : 'width:clamp(44px,11vw,80px);margin-bottom:clamp(12px,1.8vw,20px);';
    // svg_inline 에 인라인 스타일 주입 (currentColor 사용 전제)
    svg = injectSvgStyle(opt.svg_inline, `${w}height:auto;display:block;${op}`);
  }

  const en = opt.name_en
    ? `<span style="font-family:'Outfit',sans-serif;font-size:clamp(9px,1.2vw,12px);font-weight:700;color:${enCol};letter-spacing:0.15em;line-height:1;">${esc(opt.name_en)}</span>`
    : '';
  // 표시값(id) = 숫자 크게 + 단위 자동 2/3. 이름·영문 없으면 옆 칸 자체를 숨겨 여백 방지.
  const bigVal = `<span style="font-family:${FONT_EN};font-size:clamp(28px,5.5vw,64px);font-weight:900;color:${letterCol};line-height:0.9;letter-spacing:-0.02em;flex-shrink:0;">${bigValue(opt.id)}</span>`;
  const nameSpan = opt.name_ko ? `<span style="font-size:clamp(13px,1.7vw,18px);font-weight:700;color:${nameCol};line-height:1.2;">${esc(opt.name_ko)}</span>` : '';
  const nameBlock = (opt.name_ko || opt.name_en) ? `<div style="display:flex;flex-direction:column;gap:clamp(2px,0.4vw,4px);min-width:0;">${nameSpan}${en}</div>` : '';

  const head = dark
    ? `<div style="display:flex;justify-content:space-between;align-items:center;gap:clamp(6px,1vw,10px);margin-bottom:clamp(10px,1.5vw,16px);">
         <div style="display:flex;align-items:center;gap:clamp(10px,1.4vw,14px);min-width:0;">${bigVal}${nameBlock}</div>${check}
       </div>`
    : `<div style="display:flex;align-items:center;gap:clamp(10px,1.4vw,14px);margin-bottom:clamp(10px,1.5vw,16px);">${bigVal}${nameBlock}</div>`;

  return `<div style="${bg}border-radius:clamp(8px,1.2vw,12px);padding:clamp(14px,2.5vw,28px);">
    ${svg}${head}
    <div style="height:1px;background:${divCol};margin-bottom:clamp(10px,1.5vw,16px);"></div>
    <div style="font-size:clamp(10px,1.3vw,13px);color:${descCol};line-height:1.6;">${nl2br(opt.description_ko || '')}</div>
  </div>`;
}

/* svg_inline 문자열 첫 <svg ...> 에 style 주입 */
function injectSvgStyle(svgStr, style) {
  return svgStr.replace(/<svg\b([^>]*?)>/i, (m, attrs) => {
    if (/style=/.test(attrs)) {
      return m.replace(/style="([^"]*)"/i, (mm, s) => `style="${s};${style}"`);
    }
    return `<svg${attrs} style="${style}">`;
  });
}

/* 카드 그룹 (라벨 + 3열 그리드) */
function cardGroup(card, selectedId, variant) {
  if (!card) return '';
  let opts = card.options || [];
  // single_display 카드(예: 틴닝 발수/홈수/감모 — 옵션이 많음)는 상세페이지에서 '선택한 것만' 표시.
  // (NEUTRAL=작업대 정보보기 모드에서는 전체 표시)
  if (card.single_display && !NEUTRAL) {
    const chosen = opts.find(o => o.id === selectedId) || opts[0];
    opts = chosen ? [chosen] : [];
  }
  const cards = opts.map(o => optionCard(o, o.id === selectedId, variant, card)).join('');
  const cols = opts.length === 1 ? 1 : (opts.length === 2 ? 2 : 3);
  return `<div style="margin-bottom:clamp(48px,6vw,72px);">
    <div style="display:flex;align-items:baseline;gap:clamp(12px,1.5vw,16px);margin-bottom:clamp(24px,3vw,32px);flex-wrap:wrap;">
      <span style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:800;color:#1A1A1A;letter-spacing:0.15em;">${esc(card.label_ko)}</span>
      <span style="font-size:clamp(12px,1.5vw,14px);color:#8A8580;letter-spacing:0.02em;">— ${esc(card.label_subtitle_ko || '')}</span>
    </div>
    <div style="display:grid;grid-template-columns:repeat(${cols},1fr);gap:clamp(6px,1vw,12px);">${cards}</div>
  </div>`;
}

/* 틴닝 전용 — 발 · 홈 · 감모를 한 행에 나란히 + 하단 총정리(2행) */
function thinningRow(spec, catalog) {
  const cts = ['thinning_teeth', 'thinning_holes', 'thinning_reduction'];
  const cells = [];
  const summary = [];
  for (const ct of cts) {
    const c = catalog.byCardType[ct];
    if (!c || !(c.applies_to || []).includes(spec.type)) continue;
    const selId = spec.selections && spec.selections[ct];
    const opt = (c.options || []).find(o => o.id === selId) || (c.options || [])[0];
    if (!opt) continue;
    cells.push({ c, opt });
    // 총정리 = 표시값(id) + 이름(name_ko), 단위는 자동 2/3 → "24발 · 3홈 · 20%"
    summary.push(bigValue((opt.id || '') + (opt.name_ko || '')));
  }
  if (!cells.length) return '';
  const note = (spec.custom_fields && spec.custom_fields.thinning_note) || '';
  // 홈 형태 SVG — 선택 카드(thinning_shape)에서 고른 배경 없는 SVG를 스펙 위에 표시
  let shapeHtml = '';
  const shapeCard = catalog.byCardType['thinning_shape'];
  if (shapeCard) {
    const sSel = spec.selections && spec.selections['thinning_shape'];
    const sOpt = (shapeCard.options || []).find(o => o.id === sSel) || (shapeCard.options || [])[0];
    if (sOpt && sOpt.svg_inline) {
      shapeHtml = `<div style="margin-bottom:clamp(28px,3.5vw,44px);">
        <div style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:800;color:#1A1A1A;letter-spacing:0.15em;margin-bottom:clamp(16px,2vw,22px);">홈 형태 <span style="color:#8A8580;font-weight:700;">— SHAPE</span></div>
        <div style="background:#F5F3F0;border-radius:clamp(10px,1.4vw,16px);padding:clamp(28px,4vw,52px) clamp(20px,3vw,40px);color:#1A1A1A;text-align:center;">
          ${injectSvgStyle(sOpt.svg_inline, 'width:min(clamp(180px,42vw,420px),100%);height:auto;display:block;margin:0 auto;')}
        </div>
      </div>`;
    }
  }
  const _n = NEUTRAL; NEUTRAL = true;   // 요약 = 라이트 카드(선명한 다크 텍스트/SVG)
  const cardsHtml = cells.map(({ c, opt }) => `<div style="min-width:0;">
      <div style="font-family:'Outfit',sans-serif;font-size:clamp(9px,1.2vw,12px);font-weight:700;color:#8A8580;letter-spacing:0.12em;text-transform:uppercase;margin-bottom:clamp(8px,1vw,12px);">${esc(c.label_subtitle_ko || c.label_ko || '')}</div>
      ${optionCard(opt, false, 'text', c)}
    </div>`).join('');
  NEUTRAL = _n;
  return `<div style="margin-bottom:clamp(48px,6vw,72px);">
    ${shapeHtml}
    <div style="display:flex;align-items:baseline;gap:clamp(12px,1.5vw,16px);margin-bottom:clamp(20px,2.5vw,28px);flex-wrap:wrap;">
      <span style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:800;color:#1A1A1A;letter-spacing:0.15em;">THINNING SPEC</span>
      <span style="font-size:clamp(12px,1.5vw,14px);color:#8A8580;">— 발 · 홈 · 감모</span>
    </div>
    <div style="display:grid;grid-template-columns:repeat(${cells.length},1fr);gap:clamp(6px,1vw,12px);">${cardsHtml}</div>
    <div style="margin-top:clamp(12px,1.6vw,18px);padding:clamp(20px,2.8vw,32px) clamp(16px,2.2vw,24px);background:#1A1A1A;border-radius:clamp(8px,1.2vw,12px);text-align:center;">
      <span style="font-family:'Outfit',sans-serif;font-size:clamp(18px,3.4vw,32px);font-weight:800;color:#FAF9F7;letter-spacing:0.02em;">${summary.join('  ·  ')}</span>
      ${note ? `<p style="margin:clamp(14px,1.8vw,20px) auto 0;max-width:560px;font-size:clamp(13px,1.7vw,15px);color:rgba(245,245,243,0.72);line-height:1.75;">${nl2br(esc(note))}</p>` : ''}
    </div>
  </div>`;
}

/* HANDLE — grip(3 카드 세로) + camel(2 카드 가로 컴팩트) */
function handleGroup(spec, catalog) {
  const grip = catalog.byCardType['handle_grip'];
  const camel = catalog.byCardType['handle_camel'];
  if (!grip && !camel) return '';
  const selGrip = spec.selections?.handle_grip;
  const selCamel = spec.selections?.handle_camel;

  // 가로 SVG 영역(핸들 이미지) — grip/camel 공통. 카드색(currentColor) 따라 핸들 라인색 자동
  const handleBand = (svgInline, sel) => {
    const dark = sel && !NEUTRAL;
    const bandBg = dark ? 'rgba(245,245,243,0.07)' : '#F5F3F0';
    const op = (dark || NEUTRAL) ? '' : 'opacity:0.4;';
    return `<div style="background:${bandBg};border-radius:clamp(6px,1vw,10px);padding:clamp(8px,1.4vw,12px) clamp(10px,1.6vw,14px);margin-bottom:clamp(10px,1.4vw,14px);">${injectSvgStyle(svgInline, `width:100%;height:auto;display:block;${op}`)}</div>`;
  };
  // 그립 카드 = 가로 SVG 영역(밴드) + 이름 + 설명. flex 2열(+3번째 중앙) — flex-basis 47%
  const gripCard = (o, sel) => {
    const dark = sel && !NEUTRAL;
    const bg = dark ? 'background:#1A1A1A;color:#FAF9F7;' : 'background:#FFFFFF;border:1px solid #EDEBE8;';
    const nameCol = dark ? '#FAF9F7' : '#1A1A1A';
    const descCol = dark ? 'rgba(245,245,243,0.65)' : (NEUTRAL ? '#2D2D2D' : '#8A8580');
    const enCol = dark ? 'rgba(245,245,243,0.45)' : (NEUTRAL ? '#8A8580' : '#B8B4AF');
    const check = dark ? `<div style="position:absolute;top:clamp(8px,1.2vw,12px);right:clamp(10px,1.4vw,14px);font-size:clamp(11px,1.4vw,14px);color:#FAF9F7;font-weight:700;line-height:1;">✓</div>` : '';
    return `<div style="min-width:0;${bg}border-radius:clamp(8px,1.2vw,12px);padding:clamp(12px,1.8vw,18px);position:relative;box-sizing:border-box;">${check}
      ${o.svg_inline ? handleBand(o.svg_inline, sel) : ''}
      <div style="font-size:clamp(11px,1.4vw,14px);font-weight:700;color:${nameCol};text-align:center;line-height:1.3;margin-bottom:${o.name_en ? '2px' : 'clamp(4px,0.6vw,6px)'};">${esc(o.name_ko)}</div>
      ${o.name_en ? `<div style="font-family:'Outfit',sans-serif;font-size:clamp(8px,1vw,10px);font-weight:700;letter-spacing:0.12em;color:${enCol};text-align:center;line-height:1;margin-bottom:clamp(5px,0.8vw,8px);">${esc(o.name_en)}</div>` : ''}
      <div style="font-size:clamp(9px,1.1vw,11px);color:${descCol};text-align:center;line-height:1.4;">${nl2br(o.description_ko || '')}</div>
    </div>`;
  };

  // 카멜/플랫 = 아이콘 + 글씨 같은 행 (가로 컴팩트, 2차 정보 — 절제). 2열 고정
  const camelCard = (o, sel) => {
    const dark = sel && !NEUTRAL;
    const bg = dark ? 'background:#1A1A1A;color:#FAF9F7;' : 'background:#FFFFFF;border:1px solid #EDEBE8;';
    const nameCol = dark ? '#FAF9F7' : '#1A1A1A';
    const descCol = dark ? 'rgba(245,245,243,0.6)' : (NEUTRAL ? '#2D2D2D' : '#8A8580');
    const enCol = dark ? 'rgba(245,245,243,0.45)' : (NEUTRAL ? '#8A8580' : '#B8B4AF');
    const svg = o.svg_inline ? injectSvgStyle(o.svg_inline, `width:min(clamp(40px,9vw,116px),42%);height:auto;flex-shrink:0;display:block;${(dark || NEUTRAL) ? '' : 'opacity:0.4;'}`) : '';
    const check = dark ? `<div style="position:absolute;top:clamp(5px,0.7vw,8px);right:clamp(8px,1vw,12px);font-size:clamp(9px,1.2vw,12px);color:#FAF9F7;font-weight:700;line-height:1;">✓</div>` : '';
    return `<div style="${bg}border-radius:clamp(8px,1.2vw,12px);padding:clamp(10px,1.4vw,16px);display:flex;align-items:center;gap:clamp(10px,1.4vw,14px);position:relative;">${check}${svg}
      <div style="flex:1;min-width:0;">
        <div style="font-size:clamp(11px,1.3vw,13px);font-weight:700;color:${nameCol};line-height:1.2;">${esc(o.name_ko)}</div>
        ${o.name_en ? `<div style="font-family:'Outfit',sans-serif;font-size:clamp(8px,1vw,10px);font-weight:700;letter-spacing:0.1em;color:${enCol};line-height:1;margin:1px 0 3px;">${esc(o.name_en)}</div>` : '<div style="height:3px"></div>'}
        <div style="font-size:clamp(9px,1.1vw,11px);color:${descCol};line-height:1.3;">${nl2br(o.description_ko || '')}</div>
      </div>
    </div>`;
  };

  const gripCards = grip ? (grip.options || []).map(o => gripCard(o, o.id === selGrip)).join('') : '';
  const camelCards = camel ? (camel.options || []).map(o => camelCard(o, o.id === selCamel)).join('') : '';

  return `<div style="margin-bottom:clamp(48px,6vw,72px);">
    <div style="display:flex;align-items:baseline;gap:clamp(12px,1.5vw,16px);margin-bottom:clamp(24px,3vw,32px);flex-wrap:wrap;">
      <span style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:800;color:#1A1A1A;letter-spacing:0.15em;">HANDLE</span>
    </div>
    ${grip ? `<div style="margin-bottom:clamp(16px,2vw,24px);">
      <div style="font-family:'Outfit',sans-serif;font-size:clamp(10px,1.2vw,12px);color:#8A8580;font-weight:700;letter-spacing:0.15em;margin-bottom:clamp(10px,1.3vw,14px);">${esc(grip.label_ko || '')}${grip.label_subtitle_ko ? ` <span style="font-family:'Plus Jakarta Sans','Noto Sans KR',sans-serif;font-weight:500;letter-spacing:0;color:#B8B4AF;">— ${esc(grip.label_subtitle_ko)}</span>` : ''}</div>
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:clamp(8px,1.2vw,14px);">${gripCards}</div>
    </div>` : ''}
    ${camel ? `<div>
      <div style="font-family:'Outfit',sans-serif;font-size:clamp(10px,1.2vw,12px);color:#8A8580;font-weight:700;letter-spacing:0.15em;margin-bottom:clamp(8px,1vw,12px);">${esc(camel.label_ko || '')}${camel.label_subtitle_ko ? ` <span style="font-family:'Plus Jakarta Sans','Noto Sans KR',sans-serif;font-weight:500;letter-spacing:0;color:#B8B4AF;">— ${esc(camel.label_subtitle_ko)}</span>` : ''}</div>
      <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:clamp(8px,1.2vw,14px);">${camelCards}</div>
    </div>` : ''}
  </div>`;
}

/* GRIP SIZE — 일러스트 + 측정값 + 핸들 설명 */
function gripSizeBlock(spec, catalog) {
  const cf = spec.custom_fields || {};
  const thumb = cf.grip_thumb || '—';
  const ring = cf.grip_ring || '—';
  const desc = resolveCopy(spec, catalog, 'handle_description', { customKey: 'handle_description' }) || '';
  const gripSvg = `${IMG_HOST}/projects/products/shared/scissors-grip.svg`;
  return `<div style="margin-bottom:clamp(48px,6vw,72px);">
    <div style="display:flex;align-items:baseline;gap:clamp(12px,1.5vw,16px);margin-bottom:clamp(24px,3vw,32px);flex-wrap:wrap;">
      <span style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:800;color:#1A1A1A;letter-spacing:0.15em;">GRIP SIZE</span>
      <span style="font-size:clamp(12px,1.5vw,14px);color:#8A8580;letter-spacing:0.02em;">— 손가락 구멍 크기</span>
    </div>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:clamp(20px,2.5vw,40px);align-items:start;">
      <div style="background:#FFFFFF;border:1px solid #EDEBE8;border-radius:clamp(8px,1.2vw,12px);aspect-ratio:5/4;display:flex;align-items:center;justify-content:center;color:#1A1A1A;padding:clamp(20px,3vw,32px);">
        <img src="${gripSvg}" alt="가위 그립 일러스트" style="width:100%;height:auto;display:block;">
      </div>
      <div style="display:flex;flex-direction:column;gap:clamp(20px,2.5vw,32px);">
        <div>
          <div style="display:flex;justify-content:space-between;align-items:baseline;padding:clamp(14px,1.8vw,18px) 0;border-bottom:1px solid #EDEBE8;">
            <span style="font-size:clamp(13px,1.6vw,15px);color:#8A8580;font-weight:500;">엄지부</span>
            <span style="font-family:'Outfit',sans-serif;font-size:clamp(15px,1.9vw,18px);color:#1A1A1A;font-weight:700;letter-spacing:0.02em;">${esc(thumb)}<span style="font-weight:500;color:#8A8580;font-size:0.7em;margin-left:4px;">mm</span></span>
          </div>
          <div style="display:flex;justify-content:space-between;align-items:baseline;padding:clamp(14px,1.8vw,18px) 0;">
            <span style="font-size:clamp(13px,1.6vw,15px);color:#8A8580;font-weight:500;">약지부</span>
            <span style="font-family:'Outfit',sans-serif;font-size:clamp(15px,1.9vw,18px);color:#1A1A1A;font-weight:700;letter-spacing:0.02em;">${esc(ring)}<span style="font-weight:500;color:#8A8580;font-size:0.7em;margin-left:4px;">mm</span></span>
          </div>
        </div>
        <p style="font-size:clamp(13px,1.6vw,15px);color:#2D2D2D;line-height:1.75;margin:0;">${nl2br(desc)}</p>
      </div>
    </div>
  </div>`;
}

/* WEIGHT — 무게 막대 */
function weightBlock(spec) {
  const cf = spec.custom_fields || {};
  const g = spec.weight_g != null ? spec.weight_g : (cf.weight_g || '—');
  const band = cf.weight_band || '중간';
  const pctMap = { '가벼움': 18, 'LIGHT': 18, '라이트': 18, '중간': 50, '무거움': 82, 'HEAVY': 82 };
  const pct = cf.weight_percent != null ? cf.weight_percent : (pctMap[band] != null ? pctMap[band] : 50);
  const desc = cf.weight_description || '';
  return `<div>
    <div style="display:flex;align-items:baseline;gap:clamp(12px,1.5vw,16px);margin-bottom:clamp(24px,3vw,32px);flex-wrap:wrap;">
      <span style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:800;color:#1A1A1A;letter-spacing:0.15em;">WEIGHT</span>
      <span style="font-size:clamp(12px,1.5vw,14px);color:#8A8580;letter-spacing:0.02em;">— 무게감</span>
    </div>
    <div style="background:#FFFFFF;border:1px solid #EDEBE8;border-radius:clamp(8px,1.5vw,12px);padding:clamp(28px,4vw,48px);">
      <div style="display:flex;align-items:baseline;gap:clamp(8px,1.2vw,14px);margin-bottom:clamp(20px,3vw,32px);">
        <span style="font-family:'Outfit',sans-serif;font-size:clamp(40px,6vw,72px);font-weight:900;color:#1A1A1A;line-height:1;">${esc(g)}</span>
        <span style="font-size:clamp(14px,1.8vw,18px);font-weight:600;color:#8A8580;letter-spacing:0.05em;">g</span>
      </div>
      <div style="position:relative;margin-bottom:clamp(12px,1.5vw,16px);">
        <div style="height:6px;background:#EDEBE8;border-radius:3px;width:100%;"></div>
        <div style="position:absolute;top:-5px;left:${pct}%;transform:translateX(-50%);width:16px;height:16px;background:#1A1A1A;border-radius:50%;border:3px solid #FAF9F7;box-shadow:0 0 0 1px #1A1A1A;"></div>
      </div>
      <div style="display:flex;justify-content:space-between;font-family:'Outfit',sans-serif;font-size:clamp(10px,1.3vw,12px);font-weight:700;color:#8A8580;letter-spacing:0.15em;">
        <span>LIGHT</span><span style="color:#1A1A1A;">${esc(band)}</span><span>HEAVY</span>
      </div>
      ${desc ? `<div style="height:1px;background:#EDEBE8;margin:clamp(20px,2.5vw,28px) 0 clamp(16px,2vw,20px);"></div>
      <div style="font-size:clamp(12px,1.5vw,14px);color:#2D2D2D;line-height:1.6;">${nl2br(desc)}</div>` : ''}
    </div>
  </div>`;
}

/* FOR YOU 카드 (맞다 / 안 맞다) */
function forYouCard(title, items, miss) {
  const bg = miss ? 'background:#F5F3F0;' : 'background:#FFFFFF;border:1px solid #EDEBE8;';
  const titleCol = miss ? '#8A8580' : '#1A1A1A';
  const itemCol = miss ? '#8A8580' : '#2D2D2D';
  const divCol = miss ? '#D4D0CB' : '#EDEBE8';
  const rows = (items || []).map((t, i) => {
    const border = i < items.length - 1 ? `border-bottom:1px solid ${divCol};` : '';
    return `<div style="font-size:clamp(14px,1.8vw,16px);color:${itemCol};line-height:1.9;padding:clamp(8px,1vw,10px) 0;${border}">— ${esc(t)}</div>`;
  }).join('');
  return `<div style="${bg}border-radius:clamp(8px,1.5vw,12px);padding:clamp(28px,4vw,48px);margin-bottom:clamp(16px,2vw,24px);">
    <div style="font-size:clamp(13px,1.7vw,15px);font-weight:700;color:${titleCol};letter-spacing:0.02em;margin-bottom:clamp(20px,2.5vw,28px);">${esc(title)}</div>
    ${rows}
  </div>`;
}

/* SPEC 메타 row */
function specRow(label, value, last) {
  if (!value) return '';
  return `<div style="display:flex;justify-content:space-between;align-items:center;padding:clamp(14px,1.8vw,18px) 0;${last ? '' : 'border-bottom:1px solid #D4D0CB;'}font-size:clamp(12px,1.5vw,14px);">
    <span style="color:#8A8580;font-weight:500;">${esc(label)}</span><span style="color:#1A1A1A;font-weight:700;">${esc(value)}</span>
  </div>`;
}

/* GRADE 카드 (R/A/E/S, price_grade 강조) */
function gradeCards(spec, catalog) {
  const grade = catalog.byCardType['grade'];
  if (!grade) return '';
  return (grade.options || []).map(o => {
    const dark = (o.id === spec.price_grade) && !NEUTRAL;
    if (dark) {
      return `<div style="background:#1A1A1A;color:#FAF9F7;border-radius:clamp(8px,1.5vw,12px);padding:clamp(24px,3vw,32px);position:relative;">
        <div style="position:absolute;top:clamp(14px,2vw,20px);right:clamp(16px,2.2vw,22px);font-size:clamp(13px,1.7vw,16px);color:#FAF9F7;font-weight:700;line-height:1;">✓</div>
        <div style="font-family:'Outfit',sans-serif;font-size:clamp(36px,5vw,52px);font-weight:900;color:#FAF9F7;line-height:1;margin-bottom:clamp(8px,1vw,12px);">${esc(o.id)}</div>
        <div style="font-size:clamp(11px,1.4vw,13px);font-weight:700;color:rgba(245,245,243,0.55);letter-spacing:0.15em;margin-bottom:clamp(12px,1.5vw,16px);">${esc(o.name_en || '')}</div>
        <p style="font-size:clamp(13px,1.6vw,15px);color:rgba(245,245,243,0.85);line-height:1.7;margin:0;">${nl2br(o.description_ko || '')}</p>
      </div>`;
    }
    // 중립(정보) 보기 = 모든 등급 또렷 풀카드 (글자+영문+설명, ✓ 없음)
    if (NEUTRAL) {
      return `<div style="background:#FFFFFF;border:1px solid #EDEBE8;border-radius:clamp(8px,1.5vw,12px);padding:clamp(24px,3vw,32px);">
        <div style="font-family:'Outfit',sans-serif;font-size:clamp(36px,5vw,52px);font-weight:900;color:#1A1A1A;line-height:1;margin-bottom:clamp(8px,1vw,12px);">${esc(o.id)}</div>
        <div style="font-size:clamp(11px,1.4vw,13px);font-weight:700;color:#8A8580;letter-spacing:0.15em;margin-bottom:clamp(12px,1.5vw,16px);">${esc(o.name_en || '')}</div>
        <p style="font-size:clamp(13px,1.6vw,15px);color:#2D2D2D;line-height:1.7;margin:0;">${nl2br(o.description_ko || '')}</p>
      </div>`;
    }
    // 미선택 등급 = 컴팩트 (글자 + 영문 라벨만, 설명 생략) — 모델 등급만 강조, 여백 유지
    return `<div style="background:#FFFFFF;border:1px solid #EDEBE8;border-radius:clamp(8px,1.5vw,12px);padding:clamp(16px,2vw,20px) clamp(20px,2.5vw,24px);display:flex;align-items:center;gap:clamp(12px,1.6vw,16px);align-self:start;">
      <span style="font-family:'Outfit',sans-serif;font-size:clamp(26px,3.4vw,38px);font-weight:900;color:#1A1A1A;line-height:1;flex-shrink:0;">${esc(o.id)}</span>
      <span style="font-size:clamp(11px,1.4vw,13px);font-weight:700;color:#8A8580;letter-spacing:0.15em;">${esc(o.name_en || '')}</span>
    </div>`;
  }).join('');
}

/* SAME HANDLE 라인업 */
function lineupCards(spec) {
  const items = spec.lineup || [];
  const main = `<div>
    <div style="width:100%;aspect-ratio:2/3;background:#1A1A1A;color:#FAF9F7;display:flex;flex-direction:column;align-items:center;justify-content:center;font-size:10px;letter-spacing:0.05em;border-radius:4px;margin-bottom:clamp(8px,1vw,12px);text-align:center;padding:8px;position:relative;">
      <div style="position:absolute;top:8px;right:10px;font-size:clamp(11px,1.4vw,14px);color:#FAF9F7;font-weight:700;line-height:1;">✓</div>
      <div style="color:rgba(245,245,243,0.5);">[ ${esc(spec.model)}<br>본 모델 ]</div>
    </div>
    <div style="font-family:'Outfit',sans-serif;font-size:clamp(12px,1.5vw,14px);font-weight:700;color:#1A1A1A;letter-spacing:0.02em;">${esc(spec.model)}</div>
    <div style="font-size:clamp(10px,1.3vw,12px);color:#8A8580;margin-top:2px;">${esc(spec.size_inch != null ? spec.size_inch + ' inch · 본 모델' : '본 모델')}</div>
  </div>`;
  const others = items.map(it => {
    const model = typeof it === 'string' ? it : it.model;
    const sub = typeof it === 'string' ? '' : (it.sub || '');
    return `<div>
      <div style="width:100%;aspect-ratio:2/3;background:#F5F3F0;display:flex;align-items:center;justify-content:center;color:#8A8580;font-size:10px;letter-spacing:0.05em;border-radius:4px;margin-bottom:clamp(8px,1vw,12px);text-align:center;padding:8px;">[ ${esc(model)}<br>날부 2:3 ]</div>
      <div style="font-family:'Outfit',sans-serif;font-size:clamp(12px,1.5vw,14px);font-weight:700;color:#1A1A1A;letter-spacing:0.02em;">${esc(model)}</div>
      ${sub ? `<div style="font-size:clamp(10px,1.3vw,12px);color:#8A8580;margin-top:2px;">${esc(sub)}</div>` : ''}
    </div>`;
  }).join('');
  return main + others;
}

/* ─── 브랜드 고정 섹션 (모델 무관 — v10 verbatim) ─── */
/* 고객 목소리 — 제품별 실제 후기(spec.custom_fields.reviews) 큐레이션. 없으면 브랜드 상담후기 폴백 */
const VOICES_FALLBACK = [
  ['상담받고 처음으로 "가위 안 사셔도 됩니다"란 말 들었어요. 신선한 충격이었습니다.', '김○○ 원장님', '서울 · 경력 12년'],
  ['새 가위가 필요한 게 아니라 다른 형태가 답이라고 짚어주셨어요. 그동안 잘못 골라온 것 같았습니다.', '박○○ 디자이너', '부산 · 경력 8년'],
  ['평생 무료 복원수리라서 정말 한 자루 쓰는 만큼 마음이 편해요. 처음 들어본 보장입니다.', '이○○ 원장님', '대구 · 경력 15년']
];
function voicesSection(spec) {
  const cf = spec.custom_fields || {};
  const curated = Array.isArray(cf.reviews) ? cf.reviews.filter(r => r && (r.quote || '').trim()) : [];
  const useCurated = curated.length > 0;
  const items = useCurated ? curated.map(r => [r.quote, r.name || '', r.meta || '']) : VOICES_FALLBACK;
  const sub = useCurated ? '— 실제 고객 후기' : '— 제품 후기가 아닌, 상담을 통해 느낀 점';
  return `<div style="background:#EDEBE8;padding:clamp(80px,10vw,140px) clamp(24px,4vw,48px);">
  ${eyebrow('08', 'VOICES')}
  <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(20px,3vw,32px) 0;">고객 목소리</h2>
  <p style="font-size:clamp(13px,1.6vw,15px);color:#8A8580;line-height:1.7;margin:0 0 clamp(48px,6vw,72px) 0;font-style:italic;">${sub}</p>
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:clamp(16px,2vw,24px);align-items:start;">
    ${items.map(([q, n, m]) => `<div style="background:#FFFFFF;border-radius:clamp(8px,1.5vw,12px);padding:clamp(28px,4vw,40px);">
        <div style="font-family:'Outfit',sans-serif;font-size:clamp(28px,4vw,48px);font-weight:900;color:#1A1A1A;line-height:1;margin-bottom:clamp(12px,1.5vw,16px);">"</div>
        <p style="font-size:clamp(15px,1.9vw,18px);color:#1A1A1A;line-height:1.65;margin:0 0 clamp(20px,2.5vw,28px) 0;font-weight:500;">${nl2br(esc(q))}</p>
        <div style="display:flex;align-items:center;gap:clamp(10px,1.5vw,14px);">
          <div style="width:clamp(36px,5vw,44px);height:clamp(36px,5vw,44px);border-radius:50%;background:#D4D0CB;flex-shrink:0;"></div>
          <div><div style="font-size:clamp(12px,1.5vw,14px);font-weight:700;color:#1A1A1A;">${esc(n)}</div>
          <div style="font-size:clamp(10px,1.3vw,12px);color:#8A8580;margin-top:2px;">${esc(m)}</div></div>
        </div></div>`).join('')}
  </div>
</div>`;
}

const STATIC_VS = `<div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px);">
  ${eyebrow('09', 'VS GENERAL')}
  <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(48px,7vw,72px) 0;">MAMORU vs<br>일반 가위 브랜드</h2>
  <div style="border-top:2px solid #1A1A1A;">
    <div style="display:grid;grid-template-columns:1.2fr 1fr 1fr;gap:clamp(12px,1.5vw,20px);padding:clamp(16px,2.2vw,24px) 0;border-bottom:1px solid #D4D0CB;font-size:clamp(11px,1.4vw,13px);font-weight:700;color:#8A8580;letter-spacing:0.05em;text-transform:uppercase;"><span></span><span style="color:#1A1A1A;">MAMORU</span><span>일반</span></div>
    ${[['선택 방식', '손·스타일 진단 우선', '브랜드·가격 우선'],
       ['가격 정책', '즐길 수 있는 가격<br>ON·OFF 정찰제', '높은 판매가 → 할인·사은품<br>으로 구매 유도'],
       ['복원수리', '자체 기술 · 평생 무상', '외주 또는 신품 권장'],
       ['상담', '맞춤 진단 · 무료', '판매 중심 응대'],
       ['"안 사셔도 됩니다"', '합니다', '못 합니다']]
      .map(([a, b, c], i, arr) => `<div style="display:grid;grid-template-columns:1.2fr 1fr 1fr;gap:clamp(12px,1.5vw,20px);padding:clamp(16px,2.2vw,24px) 0;${i < arr.length - 1 ? 'border-bottom:1px solid #D4D0CB;' : ''}font-size:clamp(13px,1.7vw,15px);">
        <span style="color:#8A8580;font-weight:500;">${a}</span><span style="color:#1A1A1A;font-weight:700;">${b}</span><span style="color:#8A8580;">${c}</span></div>`).join('')}
  </div>
</div>`;

const STATIC_WHY = `<div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px);">
  ${eyebrow('10', 'WHY MAMORU')}
  <div style="font-family:'Outfit',sans-serif;font-size:clamp(48px,7vw,80px);font-weight:900;color:#1A1A1A;line-height:1;margin-bottom:clamp(20px,3vw,28px);">"</div>
  <p style="font-size:clamp(20px,3.2vw,36px);color:#1A1A1A;line-height:1.4;font-weight:600;letter-spacing:-0.01em;margin:0 0 clamp(40px,5vw,56px) 0;max-width:680px;">대부분의 가위는 '유행'에 맞춰져 있습니다.<br>마모루는 '당신의 손'에 맞춥니다.</p>
  <div style="display:flex;align-items:center;gap:clamp(14px,2vw,18px);margin-bottom:clamp(40px,5vw,56px);">
    <div style="width:clamp(56px,7vw,72px);height:clamp(56px,7vw,72px);border-radius:50%;background:#D4D0CB;flex-shrink:0;display:flex;align-items:center;justify-content:center;color:#8A8580;font-size:9px;letter-spacing:0.1em;">[얼굴]</div>
    <div><div style="font-size:clamp(13px,1.7vw,15px);font-weight:700;color:#1A1A1A;letter-spacing:0.02em;">백성민</div>
    <div style="font-size:clamp(11px,1.4vw,13px);color:#8A8580;letter-spacing:0.05em;margin-top:3px;">MAMORU 대표 · 컨설팅 · 복원수리 전문</div></div>
  </div>
</div>
<div style="width:100%;aspect-ratio:3/2;max-height:600px;background:#F5F3F0;display:flex;align-items:center;justify-content:center;color:#8A8580;font-size:12px;letter-spacing:0.1em;">[ IMG_workshop — 공방 / 가위 점검 사진 ]</div>
<div style="padding:clamp(40px,5vw,72px) clamp(20px,3vw,40px) clamp(80px,10vw,140px);">
  <p style="font-size:clamp(14px,1.8vw,17px);color:#2D2D2D;line-height:1.85;max-width:620px;margin:0 0 clamp(16px,2vw,20px) 0;">아버지에게 복원 기술을 물려받고, 일본 공장에서 직접 견학하며 배웠습니다. 미용가위만 10년을 만져왔습니다.</p>
  <p style="font-size:clamp(14px,1.8vw,17px);color:#2D2D2D;line-height:1.85;max-width:620px;margin:0;"><strong style="color:#1A1A1A;font-weight:700;">판매하지 않습니다. 안내할 뿐입니다.</strong> 모든 고객 동일 가격, 할인 없음, 평생 자체 복원수리.</p>
</div>`;

function staticCTA(spec) {
  const link = (spec.custom_fields && spec.custom_fields.cta_link) || '#';
  return `<div style="background:#1A1A1A;color:#FAF9F7;padding:clamp(80px,10vw,140px) clamp(24px,4vw,48px);text-align:left;">
    ${eyebrow('13', 'NEXT STEP', true)}
    <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(28px,5.5vw,72px);font-weight:800;color:#FAF9F7;letter-spacing:-0.03em;line-height:1.05;margin:0 0 clamp(28px,4vw,40px) 0;">안 사셔도<br>괜찮습니다.</h2>
    <p style="font-size:clamp(15px,1.9vw,18px);color:rgba(245,245,243,0.75);line-height:1.7;max-width:580px;margin:0 0 clamp(48px,6vw,72px) 0;">먼저 본인의 손과 스타일을 정확히 알고 싶다면, 무료 컨설팅을 통해 안내드립니다. 구매는 그 다음입니다.</p>
    <div style="max-width:480px;"><a href="${esc(link)}" style="display:block;padding:clamp(18px,2.2vw,22px) clamp(20px,3vw,28px);background:#FAF9F7;color:#1A1A1A;text-decoration:none;font-size:clamp(14px,1.8vw,16px);font-weight:700;letter-spacing:0.02em;text-align:center;border-radius:8px;">맞춤 컨설팅 신청</a></div>
    <div style="margin-top:clamp(56px,7vw,96px);padding-top:clamp(28px,4vw,40px);border-top:1px solid rgba(245,245,243,0.15);">
      <p style="font-size:clamp(10px,1.3vw,11px);color:rgba(245,245,243,0.3);margin:0;">MAMORU · 마모루미용가위</p>
    </div>
  </div>`;
}

/* ════════════════ 메인 생성 함수 ════════════════ */
function renderDetailHTML(spec, catalog) {
  const type = spec.type || 'blunt';
  const sm = spec.spec_meta || {};
  const heroSub = resolveCopy(spec, catalog, 'hero_subtitle', { customKey: 'hero_subtitle_text' });
  const honestReco = resolveCopy(spec, catalog, 'honest_reco', { customKey: 'honest_reco_text' });
  const aboutBody = resolveCopy(spec, catalog, 'about_body', { customKey: 'hero_quote_about' })
                  || resolveCopy(spec, catalog, 'about_body', { customKey: 'about_body_text' });
  const aboutQuote = resolveCopy(spec, catalog, 'about_quote', { customKey: 'about_brand_quote' })
                  || resolveCopy(spec, catalog, 'about_quote', { customKey: 'about_quote_text' });
  const matchItems = resolveCopy(spec, catalog, 'for_you_match') || [];
  const missItems = resolveCopy(spec, catalog, 'for_you_miss') || [];

  // PROFILE 카드 그룹 (종류 분기 — 적용되는 카드만)
  const design = catalog.byCardType['blade_design'];
  let profileCards = '';
  // 날 선(edge) — 종류별 카드(블런트/장가위/슬라이싱) 중 이 종류에 맞는 것만
  for (const ct of ['blade_edge', 'blade_edge_long', 'blade_edge_dry']) {
    const c = catalog.byCardType[ct];
    if (c && (c.applies_to || []).includes(type))
      profileCards += cardGroup(c, spec.selections?.[ct], 'edge');
  }
  if (design && (design.applies_to || []).includes(type))
    profileCards += cardGroup(design, spec.selections?.blade_design, 'design');
  // 틴닝: 발·홈·감모를 한 행에 나란히 + 하단 총정리 / 드라이 등은 개별 카드
  if (type === 'thinning') profileCards += thinningRow(spec, catalog);
  // dry_cutting_style 폐기(2026-07-12) — blade_edge_dry 와 내용이 겹쳐 일원화. DRY 주문 옵션은 blade_edge_dry 가 담당
  for (const ct of ['thinning_teeth', 'thinning_holes', 'thinning_reduction']) {
    if (type === 'thinning' && ct.indexOf('thinning_') === 0) continue; // 위 thinningRow가 처리
    const c = catalog.byCardType[ct];
    if (c && (c.applies_to || []).includes(type))
      profileCards += cardGroup(c, spec.selections?.[ct], 'text');
  }
  profileCards += handleGroup(spec, catalog);
  profileCards += gripSizeBlock(spec, catalog);
  profileCards += weightBlock(spec);

  // SPEC 메타 rows
  const gradeOpt = catalog.cardOption('grade', spec.price_grade);
  const gradeLabel = spec.price_grade ? `${spec.price_grade}${gradeOpt ? ' · ' + (gradeOpt.name_en || '') : ''}` : '';

  return `<div class="mamoru-detail-master" style="max-width:840px;margin:40px auto;padding:0;font-family:'Plus Jakarta Sans','Noto Sans KR',sans-serif;color:#1A1A1A;line-height:1.6;-webkit-font-smoothing:antialiased;box-sizing:border-box;background:#FAF9F7;">

  <!-- 01 Hero -->
  <div style="padding:clamp(40px,5vw,72px) clamp(20px,3vw,40px) clamp(20px,3vw,32px);">
    <div style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:700;color:#8A8580;letter-spacing:0.25em;">— ${esc(spec.category_label || '')}</div>
  </div>
  <img src="${imgURL(spec, (spec.images && spec.images.hero) || 'hero.png')}" alt="${esc(spec.model)} 메인" style="display:block;width:100%;height:auto;background:#F5F3F0;">
  <div style="padding:clamp(40px,6vw,80px) clamp(20px,3vw,40px) clamp(56px,7vw,96px);">
    <h1 style="font-family:'Paperlogy','Outfit',sans-serif;font-size:clamp(40px,10vw,112px);font-weight:900;color:#1A1A1A;letter-spacing:-0.04em;line-height:1;margin:0 0 clamp(28px,4vw,48px) 0;white-space:nowrap;">${esc(spec.model)}</h1>
    <p style="font-size:clamp(16px,2.4vw,22px);color:#4A4A4A;line-height:1.5;font-weight:300;max-width:520px;margin:0 0 clamp(24px,3vw,32px) 0;">${nl2br(heroSub)}</p>
    <div style="display:flex;gap:clamp(20px,3vw,40px);font-size:clamp(12px,1.5vw,14px);color:#8A8580;letter-spacing:0.05em;font-weight:500;flex-wrap:wrap;">
      <span>${esc(spec.size_inch)} inch</span><span style="color:#D4D0CB;">·</span>
      <span>${esc(spec.weight_g)} g</span><span style="color:#D4D0CB;">·</span>
      <span>${esc(TYPE_LABEL[type] || type)}</span>
    </div>
    ${honestReco ? `<div style="margin-top:clamp(28px,4vw,40px);padding:clamp(20px,3vw,28px);background:#F5F3F0;border-radius:clamp(10px,1.5vw,14px);border-left:3px solid #1A1A1A;max-width:560px;">
      <div style="font-family:'Outfit',sans-serif;font-size:clamp(10px,1.3vw,12px);font-weight:800;letter-spacing:0.18em;color:#8A8580;text-transform:uppercase;margin-bottom:clamp(8px,1vw,12px);">MAMORU의 솔직 추천</div>
      <p style="font-size:clamp(14px,1.9vw,17px);color:#2D2D2D;line-height:1.75;margin:0;">${nl2br(honestReco)}</p>
    </div>` : ''}
  </div>

  <!-- 02 Detail -->
  <div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px) clamp(40px,5vw,64px);">${eyebrow('01', 'DETAIL')}</div>
  <img src="${imgURL(spec, (spec.images && spec.images.blade2) || 'blade2.png')}" alt="${esc(spec.model)} 날부" style="display:block;width:100%;height:auto;background:#F5F3F0;margin-bottom:clamp(8px,1.2vw,16px);">
  <img src="${imgURL(spec, (spec.images && spec.images.handle) || 'handle.png')}" alt="${esc(spec.model)} 핸들부" style="display:block;width:100%;height:auto;background:#F5F3F0;margin-bottom:clamp(8px,1.2vw,16px);">
  <img src="${imgURL(spec, (spec.images && spec.images.back) || 'back.png')}" alt="${esc(spec.model)} 뒷면" style="display:block;width:100%;height:auto;background:#F5F3F0;margin-bottom:clamp(40px,5vw,72px);">
  <div style="padding:0 clamp(20px,3vw,40px);"><div style="font-family:'Outfit',sans-serif;font-size:clamp(11px,1.4vw,13px);font-weight:600;color:#8A8580;letter-spacing:0.18em;margin-bottom:clamp(16px,2vw,24px);">— CLOSE-UP</div></div>
  <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(360px,1fr));gap:clamp(8px,1.2vw,16px);align-items:start;padding:0 clamp(20px,3vw,40px) clamp(80px,10vw,140px);">
    <img src="${imgURL(spec, (spec.images && spec.images.bolt) || 'bolt.png')}" alt="볼트부" style="display:block;width:100%;height:auto;background:#F5F3F0;">
    <img src="${imgURL(spec, (spec.images && spec.images.model) || 'model.png')}" alt="모델명" style="display:block;width:100%;height:auto;background:#F5F3F0;">
  </div>

  <!-- 03 In Action -->
  <div style="background:#1A1A1A;color:#FAF9F7;padding:clamp(80px,10vw,140px) 0;">
    <div style="padding:0 clamp(20px,3vw,40px) clamp(40px,5vw,64px);">${eyebrow('02', 'IN ACTION', true)}
      <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(28px,5.5vw,64px);font-weight:800;color:#FAF9F7;letter-spacing:-0.03em;line-height:1.1;margin:0;">한 컷,<br>한 손에 멈춥니다.</h2>
    </div>
    <img src="${imgURL(spec, (spec.images && spec.images.cut) || 'cut.gif')}" alt="컷 동작" style="display:block;width:100%;height:auto;background:#2D2D2D;">
  </div>

  <!-- 04 About -->
  <div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px) clamp(40px,5vw,64px);">${eyebrow('03', 'ABOUT')}
    <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(40px,6vw,64px) 0;">${esc(spec.model)}의 특성</h2>
  </div>
  <img src="${imgURL(spec, (spec.images && spec.images.blade1) || 'blade1.png')}" alt="${esc(spec.model)} 날부" style="display:block;width:100%;height:auto;background:#F5F3F0;margin-bottom:clamp(40px,5vw,72px);">
  <div style="padding:0 clamp(20px,3vw,40px) clamp(80px,10vw,140px);">
    <p style="font-size:clamp(15px,1.9vw,18px);color:#2D2D2D;line-height:1.9;max-width:620px;margin:0 0 clamp(20px,2.5vw,28px) 0;">${nl2br(aboutBody)}</p>
    ${aboutQuote ? `<p style="font-size:clamp(15px,1.9vw,18px);color:#1A1A1A;line-height:1.9;max-width:620px;margin:0;font-weight:500;">"${nl2br(aboutQuote)}"</p>` : ''}
  </div>

  <!-- 05 Profile -->
  <div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px);">${eyebrow('04', 'PROFILE')}
    <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(20px,3vw,32px) 0;">이 가위의 특성</h2>
    <p style="font-size:clamp(13px,1.6vw,15px);color:#8A8580;line-height:1.7;margin:0 0 clamp(40px,5vw,56px) 0;font-style:italic;">— 날 / 핸들 / 무게로 본인 작업 스타일에 맞는 타입 확인</p>
    ${profileCards}
  </div>

  <!-- 06 For You -->
  <div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px);">${eyebrow('05', 'FOR YOU?')}
    <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(40px,6vw,64px) 0;">솔직한 선택 가이드</h2>
    ${forYouCard('이런 분에게 맞습니다', matchItems, false)}
    ${forYouCard('맞지 않을 수 있습니다', missItems, true)}
  </div>

  <!-- 07 Spec -->
  <div style="background:#F5F3F0;padding:clamp(80px,10vw,140px) clamp(24px,4vw,48px);">${eyebrow('06', 'SPEC')}
    <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(32px,5vw,56px) 0;">사양</h2>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:clamp(8px,1.5vw,16px);border-top:1px solid #1A1A1A;border-bottom:1px solid #1A1A1A;padding:clamp(28px,4vw,48px) 0;margin-bottom:clamp(40px,5vw,56px);">
      <div><div style="font-family:'Outfit',sans-serif;font-size:clamp(28px,4.5vw,52px);font-weight:900;color:#1A1A1A;line-height:0.95;letter-spacing:-0.02em;">${esc(spec.size_inch)}<span style="font-size:0.45em;color:#8A8580;font-weight:700;margin-left:0.1em;">inch</span></div>
        <div style="font-family:'Outfit',sans-serif;font-size:clamp(10px,1.2vw,12px);font-weight:700;color:#8A8580;letter-spacing:0.2em;text-transform:uppercase;margin-top:clamp(8px,1vw,12px);">길이</div></div>
      <div style="border-left:1px solid #D4D0CB;padding-left:clamp(12px,2vw,20px);"><div style="font-family:'Outfit',sans-serif;font-size:clamp(28px,4.5vw,52px);font-weight:900;color:#1A1A1A;line-height:0.95;letter-spacing:-0.02em;">${esc(spec.weight_g)}<span style="font-size:0.45em;color:#8A8580;font-weight:700;margin-left:0.1em;">g</span></div>
        <div style="font-family:'Outfit',sans-serif;font-size:clamp(10px,1.2vw,12px);font-weight:700;color:#8A8580;letter-spacing:0.2em;text-transform:uppercase;margin-top:clamp(8px,1vw,12px);">무게</div></div>
      <div style="border-left:1px solid #D4D0CB;padding-left:clamp(12px,2vw,20px);"><div style="font-family:'Outfit',sans-serif;font-size:clamp(20px,3vw,32px);font-weight:800;color:#1A1A1A;line-height:1;letter-spacing:-0.01em;">${esc(TYPE_LABEL[type] || type)}</div>
        <div style="font-family:'Outfit',sans-serif;font-size:clamp(10px,1.2vw,12px);font-weight:700;color:#8A8580;letter-spacing:0.2em;text-transform:uppercase;margin-top:clamp(8px,1vw,12px);">종류</div></div>
    </div>
    ${specRow('소재', sm.material)}
    ${specRow('베어링', sm.bearing)}
    ${specRow('등급', gradeLabel)}
    ${specRow('핸들', sm.handle_label)}
    ${specRow('날선 (Edge)', sm.edge_label)}
    ${specRow('날등 (Blade)', sm.design_label, true)}
  </div>

  <!-- 08 Same Handle -->
  <div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px);">${eyebrow('07', 'SAME HANDLE')}
    <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(20px,3vw,32px) 0;">${esc(spec.lineup_title || '동일 핸들')}<br>라인업</h2>
    <p style="font-size:clamp(13px,1.6vw,15px);color:#8A8580;line-height:1.7;margin:0 0 clamp(48px,6vw,72px) 0;font-style:italic;">— 한 핸들로 가위 종류 통일. 본 모델 외 동일 라인업의 다른 가위들.</p>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:clamp(8px,1.2vw,16px);">${lineupCards(spec)}</div>
  </div>

  ${voicesSection(spec)}
  ${STATIC_VS}
  ${STATIC_WHY}

  <!-- 12 Grade -->
  <div style="padding:clamp(80px,10vw,140px) clamp(20px,3vw,40px);">${eyebrow('12', 'GRADE')}
    <h2 style="font-family:'Outfit','Plus Jakarta Sans',sans-serif;font-size:clamp(24px,4.5vw,52px);font-weight:800;color:#1A1A1A;letter-spacing:-0.02em;line-height:1.15;margin:0 0 clamp(20px,3vw,32px) 0;">${esc((catalog.byCardType['grade'] || {}).label_ko || 'MAMORU LINE UP')}</h2>
    <p style="font-size:clamp(13px,1.6vw,15px);color:#8A8580;line-height:1.7;margin:0 0 clamp(48px,6vw,72px) 0;font-style:italic;">— ${esc((catalog.byCardType['grade'] || {}).label_subtitle_ko || '경력에 따른 추천 가격 & 레벨')}</p>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:clamp(12px,1.5vw,20px);">${gradeCards(spec, catalog)}</div>
  </div>

  ${staticCTA(spec)}
</div>`;
}

if (typeof module !== 'undefined') module.exports = { renderDetailHTML, esc };
