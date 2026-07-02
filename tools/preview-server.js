/* ──────────────────────────────────────────────────────────────
   로컬 실시간 미리보기 서버 (의존성 0, 라이브리로드)
   - 클로드가 projects/ 안 파일을 수정하면 → 열려있는 브라우저가 자동 새로고침
   - push / GitHub Pages 빌드 대기 없이 즉시 확인 (배포는 다 정한 뒤 한 번에)
   사용: 저장소 루트의 preview.bat 더블클릭 (또는 `node tools/preview-server.js`)
   ────────────────────────────────────────────────────────────── */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const PORT = 8123;
const OPEN_PATH = '/projects/products/builder/index.html'; // 시작 시 자동으로 열 페이지 (빌더 = 홈)

const MIME = {
  '.html': 'text/html;charset=utf-8', '.js': 'text/javascript', '.json': 'application/json',
  '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp', '.mp4': 'video/mp4',
  '.woff2': 'font/woff2', '.woff': 'font/woff', '.ico': 'image/x-icon',
};

const clients = []; // SSE 연결 (열린 미리보기 탭들)
const RELOAD_SNIPPET =
  "<script>(function(){try{var s=new EventSource('/__reload');s.onmessage=function(){location.reload();};}catch(e){}})();</script>";

const server = http.createServer((req, res) => {
  const url = decodeURIComponent(req.url.split('?')[0]);

  // 라이브리로드 채널 (SSE)
  if (url === '/__reload') {
    res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-cache', 'Connection': 'keep-alive' });
    res.write(': connected\n\n');
    clients.push(res);
    req.on('close', () => { const i = clients.indexOf(res); if (i >= 0) clients.splice(i, 1); });
    return;
  }

  // 편집기 저장 (로컬 전용) — projects/products/catalog 안 파일만 허용 + JSON 유효성 검사
  if (req.method === 'POST' && url === '/__save') {
    let body = '';
    req.on('data', c => { body += c; if (body.length > 5e6) req.destroy(); });
    req.on('end', () => {
      try {
        const { file, content } = JSON.parse(body);
        const full = path.normalize(path.join(ROOT, file));
        const allowed = path.join(ROOT, 'projects', 'products', 'catalog');
        if (!full.startsWith(allowed)) { res.writeHead(403); res.end('{"error":"허용 경로 아님"}'); return; }
        JSON.parse(content); // 깨진 JSON 이면 throw → 저장 안 함
        fs.writeFileSync(full, content);
        res.writeHead(200, { 'Content-Type': 'application/json' }); res.end('{"ok":true}');
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json' }); res.end(JSON.stringify({ error: String(e.message) }));
      }
    });
    return;
  }

  let f = path.join(ROOT, url);
  if (url.endsWith('/') || !path.extname(f)) f = path.join(f, 'index.html');

  fs.readFile(f, (e, data) => {
    if (e) { res.writeHead(404, { 'Content-Type': 'text/plain;charset=utf-8' }); res.end('404: ' + url); return; }
    const ext = path.extname(f).toLowerCase();
    let body = data;
    if (ext === '.html') body = Buffer.from(data.toString('utf8').replace('</body>', RELOAD_SNIPPET + '</body>'), 'utf8');
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream', 'Cache-Control': 'no-store' });
    res.end(body);
  });
});

// projects/ 변경 감지 → 모든 미리보기 탭 새로고침 (디바운스)
let timer = null;
try {
  fs.watch(path.join(ROOT, 'projects'), { recursive: true }, () => {
    clearTimeout(timer);
    timer = setTimeout(() => { clients.forEach(c => { try { c.write('data: reload\n\n'); } catch (e) {} }); }, 150);
  });
} catch (e) {
  console.log('파일 감시 실패(수동 새로고침 필요):', e.message);
}

server.listen(PORT, () => {
  const u = 'http://localhost:' + PORT + OPEN_PATH;
  console.log('\n  ▶ 실시간 미리보기 실행 중');
  console.log('    ' + u);
  console.log('    (클로드가 파일 고치면 이 창의 브라우저가 자동 새로고침됩니다. 끄려면 이 검은 창 닫기)\n');
  exec('start "" "' + u + '"'); // Windows 브라우저 자동 열기
});
