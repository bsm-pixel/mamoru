// 화면 점검용 가짜 수파베이스 (2026-11-16)
//
// 화면 점검(npm run check:screen)은 TMS 를 이 가짜 서버에 붙여서 띄운다.
// → 운영 자료를 읽지도 쓰지도 않고, 로그인도 가짜 계정으로 통과한다. 미들웨어·화면 코드는 한 줄도 바꾸지 않는다.
//
// 흉내 내는 것: 로그인 확인(/auth/v1/user), 표 읽기(/rest/v1/표), 함수(/rest/v1/rpc/이름).
// 쓰기 요청은 받기만 하고 아무것도 하지 않는다(무엇이 왔는지만 기록).
import http from 'node:http';

export const FAKE_USER = {
  id: '00000000-0000-4000-8000-000000000001',
  aud: 'authenticated',
  role: 'authenticated',
  email: 'check@example.test',
  app_metadata: { provider: 'email' },
  user_metadata: {},
  created_at: '2026-01-01T00:00:00Z',
};

const b64 = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');

/** 브라우저에 심을 로그인 쿠키 값 (가짜 서버에서만 통한다) */
export function fakeSessionCookie() {
  const exp = Math.floor(Date.now() / 1000) + 3600;
  const access = `${b64({ alg: 'HS256', typ: 'JWT' })}.${b64({ sub: FAKE_USER.id, role: 'authenticated', email: FAKE_USER.email, exp })}.check`;
  const session = { access_token: access, refresh_token: 'check-refresh', token_type: 'bearer', expires_in: 3600, expires_at: exp, user: FAKE_USER };
  return 'base64-' + Buffer.from(JSON.stringify(session)).toString('base64url');
}

function filterRows(rows, params) {
  let out = rows;
  for (const [col, raw] of params) {
    if (col === 'or') {
      // 이름 검색: or=(name.ilike.%가%,company_name.ilike.%가%) — 전부 ilike 일 때만 거른다(다른 조건이 섞인 or 는 무시)
      const parts = raw.replace(/^\(|\)$/g, '').split(',').map((c) => /^(\w+)\.ilike\.%(.*)%$/.exec(c));
      if (parts.length && parts.every(Boolean)) out = out.filter((r) => parts.some(([, c, v]) => String(r[c] ?? '').toLowerCase().includes(v.toLowerCase())));
      continue;
    }
    if (['select', 'order', 'limit', 'offset', 'and', 'on_conflict', 'columns'].includes(col)) continue;
    const m = /^(eq|in|is)\.(.*)$/s.exec(raw);
    if (!m) continue;
    const [, op, val] = m;
    if (!out.some((r) => col in r)) continue; // 자료에 없는 칸으로 거르면 전부 사라진다 — 무시
    if (op === 'eq') out = out.filter((r) => String(r[col]) === val);
    else if (op === 'in') { const set = val.replace(/^\(|\)$/g, '').split(',').map((v) => v.replace(/^"|"$/g, '')); out = out.filter((r) => set.includes(String(r[col]))); }
    else if (op === 'is' && val === 'null') out = out.filter((r) => r[col] == null);
  }
  return out;
}

/**
 * @param {{ tables?: Record<string, object[]>, rpc?: Record<string, unknown> }} fixtures
 * @returns {Promise<{ url: string, writes: string[], close: () => Promise<void> }>}
 */
export function startFakeSupabase(fixtures = {}, port = 3124) {
  const tables = fixtures.tables || {};
  const rpc = fixtures.rpc || {};
  const writes = [];

  const server = http.createServer((req, res) => {
    const url = new URL(req.url, `http://127.0.0.1:${port}`);
    const cors = {
      'Access-Control-Allow-Origin': req.headers.origin || '*',
      'Access-Control-Allow-Credentials': 'true',
      'Access-Control-Allow-Headers': req.headers['access-control-request-headers'] || '*',
      'Access-Control-Allow-Methods': 'GET,POST,PATCH,PUT,DELETE,HEAD,OPTIONS',
      'Access-Control-Expose-Headers': 'Content-Range',
    };
    const send = (status, body, extra = {}) => {
      res.writeHead(status, { 'Content-Type': 'application/json', ...cors, ...extra });
      res.end(req.method === 'HEAD' || body === undefined ? undefined : JSON.stringify(body));
    };
    if (req.method === 'OPTIONS') return send(204);

    req.resume(); // 보낸 내용은 읽지 않고 흘려보낸다
    req.on('end', () => {
      const p = url.pathname;
      if (p === '/auth/v1/user') return send(200, FAKE_USER);
      if (p === '/auth/v1/token') {
        const exp = Math.floor(Date.now() / 1000) + 3600;
        return send(200, { access_token: 'check.check.check', refresh_token: 'check-refresh', token_type: 'bearer', expires_in: 3600, expires_at: exp, user: FAKE_USER });
      }
      if (p.startsWith('/auth/v1/')) return send(200, {});

      if (p.startsWith('/rest/v1/rpc/')) {
        const name = p.slice('/rest/v1/rpc/'.length);
        return send(200, name in rpc ? rpc[name] : []);
      }
      if (p.startsWith('/rest/v1/')) {
        const table = p.slice('/rest/v1/'.length);
        if (req.method !== 'GET' && req.method !== 'HEAD') {
          writes.push(`${req.method} ${table}`);
          return send(200, []);
        }
        const rows = filterRows(tables[table] || [], url.searchParams);
        const range = { 'Content-Range': rows.length ? `0-${rows.length - 1}/${rows.length}` : '*/0' };
        if ((req.headers.accept || '').includes('vnd.pgrst.object')) {
          if (rows.length === 0) return send(406, { code: 'PGRST116', message: 'JSON object requested, multiple (or no) rows returned', details: 'The result contains 0 rows', hint: null }, range);
          return send(200, rows[0], range);
        }
        return send(200, rows, range);
      }
      // 저장소·실시간 등 나머지는 없는 것으로
      return send(404, { message: 'not found (check server)' });
    });
  });

  return new Promise((resolve, reject) => {
    server.on('error', reject);
    server.listen(port, '127.0.0.1', () => resolve({
      url: `http://127.0.0.1:${port}`,
      writes,
      close: () => new Promise((r) => { server.closeAllConnections?.(); server.close(() => r()); }),
    }));
  });
}
