// Make 알림톡 시나리오 → 「신호(template)마다 템플릿 칸이 읽는 값」 표를 만든다 (2026-11-22)
//
// 쓰는 법: node tests/alimtalk/snapshot-make.mjs
//   - Make 토큰은 저장소 밖(C:/마모루드라이브/ClaudeSecrets/make-token.env)에서 읽는다. 결과 파일에 토큰 · 발신 번호 · 채널 id 는 넣지 않는다.
//   - 결과: tests/alimtalk/make-reads.json — 자동 점검(tests/alimtalk/vars.test.ts)이 이 표와 TMS 코드를 맞춰 본다.
//   - Make 시나리오를 고치면 이 스크립트를 다시 돌려 표를 새로 만든다(점검이 실패하면 그게 신호다).
import fs from 'node:fs';
import path from 'node:path';

const OUT = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\//, '')), 'make-reads.json');
const env = fs.readFileSync('C:/마모루드라이브/ClaudeSecrets/make-token.env', 'utf8');
const TOKEN = (env.match(/MAKE_TOKEN=(\S+)/) || [])[1];
if (!TOKEN) throw new Error('Make 토큰을 못 찾음');
const BASE = 'https://eu2.make.com/api/v2';
const H = { Authorization: `Token ${TOKEN}` };
const TEAM = 1942714;

const list = (await (await fetch(`${BASE}/scenarios?teamId=${TEAM}`, { headers: H })).json()).scenarios;
const out = { generatedAt: new Date().toISOString(), scenarios: {} };

for (const sc of list) {
  const bp = (await (await fetch(`${BASE}/scenarios/${sc.id}/blueprint`, { headers: H })).json()).response?.blueprint;
  if (!bp) continue;
  const hook = bp.flow.find((m) => m.module === 'gateway:CustomWebHook');
  const router = bp.flow.find((m) => m.module === 'builtin:BasicRouter');
  if (!hook || !router) continue;
  const routes = [];
  for (const rt of router.routes) {
    const mod = rt.flow.find((m) => m.module === 'solapi:ata');
    if (!mod) continue;
    // 분기 조건: {{1.template}} text:equal '<신호>' (여러 줄이면 모두)
    const signals = [];
    for (const or of mod.filter?.conditions || []) for (const c of or) if (/\{\{1\.template\}\}/.test(c.a) && c.o === 'text:equal') signals.push(c.b);
    const vars = {};
    for (const [k, v] of Object.entries(mod.mapper?.variables || {})) {
      // 칸이 읽는 값 이름 — {{1.xxx}} 만 뽑는다(고정 글자 · 함수는 'fixed')
      const reads = [...String(v).matchAll(/\{\{1\.([A-Za-z0-9_]+)\}\}/g)].map((m) => m[1]);
      vars[k] = reads.length ? reads : ['(고정)'];
    }
    routes.push({ branch: mod.metadata?.designer?.name || mod.filter?.name || '', signals, templateId: mod.mapper?.templateId || '', vars });
  }
  out.scenarios[sc.id] = { name: sc.name, active: sc.isActive, routes };
}
fs.writeFileSync(OUT, JSON.stringify(out, null, 2) + '\n');
const n = Object.values(out.scenarios).reduce((s, x) => s + x.routes.length, 0);
console.log(`시나리오 ${Object.keys(out.scenarios).length}개 · 분기 ${n}개 → ${OUT}`);
