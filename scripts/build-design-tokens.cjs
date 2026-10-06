// MAMORU 디자인 시스템 tokens.json → CSS 변수 파일 생성
//   원본(SSOT): C:\마모루드라이브\마모루_클로드킷\MAMORU_디자인시스템\tokens.json  (값은 여기서만 고친다)
//   출력: projects/brand/css/mamoru-tokens.css  — 직접 수정 금지, 이 스크립트로 다시 만든다
//   실행: node scripts/build-design-tokens.cjs
const fs = require('fs');
const path = require('path');
const SRC = process.env.MAMORU_DS_DIR || 'C:/마모루드라이브/마모루_클로드킷/MAMORU_디자인시스템';
const t = JSON.parse(fs.readFileSync(path.join(SRC, 'tokens.json'), 'utf8'));

const ref = (v) => String(v).replace(/^\{(.+)\}$/, 'var(--$1)');
const base = [], light = [], dark = [];
for (const c of t.color.tokens) {
  if (typeof c.value === 'string') base.push(`  --${c.name}: ${c.value};`);
  else { light.push(`  --${c.name}: ${ref(c.value.light)};`); dark.push(`  --${c.name}: ${ref(c.value.dark)};`); }
}
const fam = Object.entries(t.type.families).map(([k, v]) => `  --font-${k}: ${v};`);
const space = t.spacing.tokens.map((s) => `  --${s.name}: ${s.value};`);
const radius = t.radius.tokens.map((r) => `  --${r.name}: ${r.value};`);
const typeCls = t.type.groups.flatMap((g) => g.styles.map((s) =>
  `.mm-${s.name}{font-family:var(--font-${g.family});font-size:${s.fontSize};line-height:${s.lineHeight};font-weight:${s.fontWeight};${s.letterSpacing ? `letter-spacing:${s.letterSpacing};` : ''}${s.name === 'label-en' ? 'text-transform:uppercase;' : ''}}`));

const css = `/* MAMORU 디자인 토큰 — 자동 생성 파일 (직접 수정 금지)
 * 원본: 마모루_클로드킷/MAMORU_디자인시스템/tokens.json (v${t.version})
 * 다시 만들기: node scripts/build-design-tokens.cjs
 * 규칙: 무채색 10색이 전부(포인트 컬러 없음) · 그림자 금지(경계선 --line 과 명도 차이) · 그라데이션 금지
 *       다크는 몰입 구간에만 — 해당 구간에 data-theme="dark" */
:root {
  /* 무채색 10색 */
${base.join('\n')}
  /* 의미 토큰 — 라이트(기본) */
${light.join('\n')}
  /* 서체 */
${fam.join('\n')}
  /* 여백 */
${space.join('\n')}
  /* 모서리 */
${radius.join('\n')}
}
/* 다크 — 몰입 구간(복원수리 안내 · 브랜드 철학)에만 */
[data-theme="dark"] {
${dark.join('\n')}
}
/* 글자 스타일 */
${typeCls.join('\n')}
`;
const out = path.join(__dirname, '..', 'projects', 'brand', 'css', 'mamoru-tokens.css');
fs.writeFileSync(out, css);
const comp = path.join(SRC, 'components.css');
if (fs.existsSync(comp)) fs.copyFileSync(comp, path.join(path.dirname(out), 'mamoru-components.css'));
console.log('written', out, css.length, 'bytes');
