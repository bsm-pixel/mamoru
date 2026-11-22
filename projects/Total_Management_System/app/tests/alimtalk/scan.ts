/**
 * 알림톡 점검 — TMS 코드가 「신호(template)마다 어떤 값을 보내는지」를 소스에서 읽어 낸다 (2026-11-22)
 *
 * sendNotification({ template, data: { … } }) 를 부르는 곳을 전부 찾아,
 *   - template: 문자열 그대로 / 조건식('a' : 'b') / 같은 파일의 변수 · 함수가 돌려주는 문자열까지 따라간다
 *   - data: 객체에 적힌 칸 이름. `...다른값` 이 섞이면 그 칸들은 알 수 없음(open) 으로 표시한다
 * 실제로 보내지 않는다 — 글자로만 읽는다.
 */
import fs from 'node:fs';
import path from 'node:path';
import ts from 'typescript';

export interface CallSite {
  file: string;          // src 기준 경로
  line: number;
  templates: string[];   // 이 곳이 보낼 수 있는 신호들
  dynamicTemplate: boolean; // 신호를 끝까지 못 따라간 경우
  keys: string[];        // data 에 적힌 칸
  open: boolean;         // ...spread 등으로 칸을 다 알 수 없음
  emptyLiterals: string[]; // 값이 '' 로 고정된 칸
}

// Make 로 가는 본문에 data 와 별개로 항상 들어가는 칸 (lib/notification/make-webhook.ts)
export const ALWAYS_SENT = ['template', 'event', 'name', 'phone', 'channel', 'sms_fallback', 'topic', '_meta'];

function walk(dir: string, out: string[] = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, out);
    else if (/\.(ts|tsx)$/.test(e.name)) out.push(p);
  }
  return out;
}

function stringsIn(node: ts.Node, sf: ts.SourceFile, seen = new Set<ts.Node>()): { values: string[]; unresolved: boolean } {
  if (seen.has(node)) return { values: [], unresolved: false };
  seen.add(node);
  const values: string[] = [];
  let unresolved = false;
  const visit = (n: ts.Node) => {
    if (ts.isStringLiteral(n) || ts.isNoSubstitutionTemplateLiteral(n)) { values.push(n.text); return; }
    if (ts.isIdentifier(n) && n !== node) {
      const r = resolveIdentifier(n, sf, seen);
      values.push(...r.values); unresolved ||= r.unresolved; return;
    }
    if (ts.isCallExpression(n) && ts.isIdentifier(n.expression)) {
      const r = resolveFunctionReturns(n.expression.text, sf, seen);
      values.push(...r.values); unresolved ||= r.unresolved; return;
    }
    // 비교식(a === 'x' ? …)의 왼쪽 비교값은 신호가 아니다 — 조건부의 결과 쪽만 본다
    if (ts.isConditionalExpression(n)) { visit(n.whenTrue); visit(n.whenFalse); return; }
    if (ts.isPropertyAccessExpression(n) || ts.isElementAccessExpression(n)) { unresolved = true; return; }
    ts.forEachChild(n, visit);
  };
  if (ts.isIdentifier(node)) {
    const r = resolveIdentifier(node, sf, seen);
    return r;
  }
  visit(node);
  return { values, unresolved };
}

function resolveIdentifier(id: ts.Identifier, sf: ts.SourceFile, seen: Set<ts.Node>) {
  const name = id.text;
  const values: string[] = [];
  let unresolved = false;
  let found = false;
  const visit = (n: ts.Node) => {
    if (ts.isVariableDeclaration(n) && ts.isIdentifier(n.name) && n.name.text === name && n.initializer) {
      found = true; const r = stringsIn(n.initializer, sf, seen); values.push(...r.values); unresolved ||= r.unresolved;
    }
    if (ts.isBinaryExpression(n) && n.operatorToken.kind === ts.SyntaxKind.EqualsToken && ts.isIdentifier(n.left) && n.left.text === name) {
      found = true; const r = stringsIn(n.right, sf, seen); values.push(...r.values); unresolved ||= r.unresolved;
    }
    // 함수 인자로 들어온 경우(template: string) — 같은 파일에서 끝까지 못 따라간다
    if (ts.isParameter(n) && ts.isIdentifier(n.name) && n.name.text === name) { found = true; unresolved = true; }
    if (ts.isBindingElement(n) && ts.isIdentifier(n.name) && n.name.text === name) { found = true; unresolved = true; }
    ts.forEachChild(n, visit);
  };
  visit(sf);
  return { values, unresolved: unresolved || !found };
}

function resolveFunctionReturns(fname: string, sf: ts.SourceFile, seen: Set<ts.Node>) {
  const values: string[] = [];
  let unresolved = false;
  let found = false;
  const visit = (n: ts.Node) => {
    if ((ts.isFunctionDeclaration(n) && n.name?.text === fname) || (ts.isVariableDeclaration(n) && ts.isIdentifier(n.name) && n.name.text === fname && n.initializer && (ts.isArrowFunction(n.initializer) || ts.isFunctionExpression(n.initializer)))) {
      found = true;
      const body = ts.isFunctionDeclaration(n) ? n : (n as ts.VariableDeclaration).initializer!;
      // 함수 안의 모든 문자열(맵 값 포함) — 넉넉하게 잡는다(신호 후보)
      const r = stringsIn(body, sf, seen); values.push(...r.values); unresolved ||= r.unresolved;
    }
    ts.forEachChild(n, visit);
  };
  visit(sf);
  return { values, unresolved: unresolved || !found };
}

export function scanCallSites(srcDir: string): CallSite[] {
  const sites: CallSite[] = [];
  for (const file of walk(srcDir)) {
    const text = fs.readFileSync(file, 'utf8');
    if (!text.includes('sendNotification(')) continue;
    const sf = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, file.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS);
    const visit = (n: ts.Node) => {
      if (ts.isCallExpression(n)) {
        const callee = n.expression.getText(sf);
        if (/(^|\.)sendNotification$/.test(callee) && n.arguments[0] && ts.isObjectLiteralExpression(n.arguments[0])) {
          const obj = n.arguments[0];
          const prop = (k: string) => obj.properties.find((p) => (ts.isPropertyAssignment(p) || ts.isShorthandPropertyAssignment(p)) && p.name.getText(sf) === k);
          const tp = prop('template');
          let templates: string[] = []; let dynamicTemplate = true;
          if (tp) {
            const expr = ts.isPropertyAssignment(tp) ? tp.initializer : (tp as ts.ShorthandPropertyAssignment).name;
            const r = stringsIn(expr, sf);
            templates = [...new Set(r.values)]; dynamicTemplate = r.unresolved;
          }
          const dp = prop('data');
          const keys: string[] = []; const emptyLiterals: string[] = []; let open = false;
          if (dp && ts.isPropertyAssignment(dp) && ts.isObjectLiteralExpression(dp.initializer)) {
            const readObj = (o: ts.ObjectLiteralExpression) => {
              for (const p of o.properties) {
                if (ts.isSpreadAssignment(p)) {
                  // ...(조건 ? { a } : {}) · ...{ a } 는 칸을 읽고, 그 밖(변수 · 함수 결과)은 알 수 없음
                  let e: ts.Expression = p.expression;
                  while (ts.isParenthesizedExpression(e)) e = e.expression;
                  const objs = ts.isObjectLiteralExpression(e) ? [e]
                    : ts.isConditionalExpression(e) ? [e.whenTrue, e.whenFalse].map((x) => { while (ts.isParenthesizedExpression(x)) x = x.expression; return x; }) : [];
                  if (objs.length && objs.every((x) => ts.isObjectLiteralExpression(x))) objs.forEach((x) => readObj(x as ts.ObjectLiteralExpression));
                  else open = true;
                  continue;
                }
                if (ts.isPropertyAssignment(p) || ts.isShorthandPropertyAssignment(p)) {
                  if (ts.isComputedPropertyName(p.name)) { keys.push(...stringsIn(p.name.expression, sf).values); continue; }
                  const k = p.name.getText(sf).replace(/^['"]|['"]$/g, '');
                  keys.push(k);
                  if (ts.isPropertyAssignment(p) && (ts.isStringLiteral(p.initializer) || ts.isNoSubstitutionTemplateLiteral(p.initializer)) && p.initializer.text === '') emptyLiterals.push(k);
                }
              }
            };
            readObj(dp.initializer);
          } else if (dp) open = true;
          const { line } = sf.getLineAndCharacterOfPosition(n.getStart(sf));
          sites.push({ file: path.relative(srcDir, file).replace(/\\/g, '/'), line: line + 1, templates, dynamicTemplate, keys, open, emptyLiterals });
        }
      }
      ts.forEachChild(n, visit);
    };
    visit(sf);
  }
  return sites;
}

/** lib/notification/make-webhook.ts 의 NotifyTemplate 에 적힌 신호 이름 전부 — 상태값 같은 다른 문자열을 신호로 잘못 잡지 않게 거른다 */
export function knownTemplates(srcDir: string): Set<string> {
  const text = fs.readFileSync(path.join(srcDir, 'lib/notification/make-webhook.ts'), 'utf8');
  const m = text.match(/export type NotifyTemplate\s*=([\s\S]*?);\s*\n/);
  return new Set([...(m?.[1] || '').matchAll(/'([a-z0-9_]+)'/g)].map((x) => x[1]));
}

export interface MakeRoute { branch: string; signals: string[]; templateId: string; vars: Record<string, string[]> }
export interface MakeReads { scenarios: Record<string, { name: string; active: boolean; routes: MakeRoute[] }> }

/** 신호마다 Make 가 읽는 값 — 같은 신호의 분기가 여럿이면(예: 수리 접수 4갈래) 모든 분기가 읽는 것(필수)과 일부만 읽는 것을 나눈다 */
export function makeNeeds(reads: MakeReads) {
  const bySignal = new Map<string, { scenario: string; branches: string[]; all: Set<string>; some: Set<string> }>();
  for (const sc of Object.values(reads.scenarios)) {
    for (const r of sc.routes) {
      const need = new Set(Object.values(r.vars).flat().filter((v) => v !== '(고정)'));
      for (const s of r.signals) {
        const cur = bySignal.get(s);
        if (!cur) { bySignal.set(s, { scenario: sc.name, branches: [r.branch], all: new Set(need), some: new Set(need) }); continue; }
        cur.branches.push(r.branch);
        for (const k of [...cur.all]) if (!need.has(k)) cur.all.delete(k);
        for (const k of need) cur.some.add(k);
      }
    }
  }
  return bySignal;
}
