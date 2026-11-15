# TMS 자료 전체 백업 — 모든 표를 날짜 폴더에 파일로 내려받는다 (읽기만 한다)
# 실행: 앱 폴더에서  PYTHONIOENCODING=utf-8 py scripts/backup-db.py [저장폴더]
#   기본 저장 위치: C:\마모루드라이브\TMS백업\YYYY-MM-DD_HHMM  (구글 드라이브로 같이 올라감)
#   30일 지난 폴더는 지운다. 사진·영수증 같은 파일 저장소는 포함하지 않는다(표 자료만).
import json, os, re, sys, shutil, urllib.request, urllib.error, datetime, hashlib

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
env = {}
for l in open(os.path.join(HERE, '.env.local'), encoding='utf-8'):
    m = re.match(r'^([A-Z0-9_]+)=(.*)$', l.strip())
    if m: env[m.group(1)] = m.group(2).strip().strip('"')
U = env['NEXT_PUBLIC_SUPABASE_URL']; S = env['SUPABASE_SERVICE_ROLE_KEY']
ROOT = sys.argv[1] if len(sys.argv) > 1 else r'C:\마모루드라이브\TMS백업'
KEEP_DAYS = 30
PAGE = 1000

def get(path, extra=None):
    h = {'apikey': S, 'Authorization': 'Bearer ' + S}
    if extra: h.update(extra)
    r = urllib.request.urlopen(urllib.request.Request(U + path, headers=h), timeout=60)
    return r.read(), dict(r.headers)

spec = json.loads(get('/rest/v1/')[0])
defs = spec.get('definitions', {})
now = datetime.datetime.now()
out = os.path.join(ROOT, now.strftime('%Y-%m-%d_%H%M'))
os.makedirs(out, exist_ok=True)

manifest = {'made_at': now.isoformat(timespec='seconds'), 'tables': {}}
problems = []
for t in sorted(defs):
    cols = list(defs[t].get('properties', {}).keys())
    order = '&order=id' if 'id' in cols else ('&order=' + cols[0] if cols else '')
    rows, off, total = [], 0, None
    try:
        while True:
            body, h = get(f'/rest/v1/{t}?select=*{order}&limit={PAGE}&offset={off}', {'Prefer': 'count=exact'})
            cr = h.get('Content-Range') or h.get('content-range') or ''
            if '/' in cr and cr.split('/')[-1].isdigit(): total = int(cr.split('/')[-1])
            chunk = json.loads(body)
            rows += chunk
            if len(chunk) < PAGE: break
            off += PAGE
    except urllib.error.HTTPError as e:
        problems.append(f'{t}: 읽기 실패 {e.code}'); continue
    raw = json.dumps(rows, ensure_ascii=False).encode('utf-8')
    with open(os.path.join(out, t + '.json'), 'wb') as f: f.write(raw)
    # 되읽기 검증: 파일을 다시 열어 건수가 서버 건수와 같은지
    back = len(json.load(open(os.path.join(out, t + '.json'), encoding='utf-8')))
    if total is not None and back != total: problems.append(f'{t}: 서버 {total}건 ≠ 파일 {back}건')
    manifest['tables'][t] = {'rows': back, 'server': total, 'sha256': hashlib.sha256(raw).hexdigest()[:16]}

manifest['problems'] = problems
with open(os.path.join(out, '_목록.json'), 'w', encoding='utf-8') as f: json.dump(manifest, f, ensure_ascii=False, indent=1)

# 오래된 폴더 정리
removed = 0
for d in os.listdir(ROOT):
    m = re.match(r'^(\d{4}-\d{2}-\d{2})_\d{4}$', d)
    if m and (now - datetime.datetime.strptime(m.group(1), '%Y-%m-%d')).days > KEEP_DAYS:
        shutil.rmtree(os.path.join(ROOT, d), ignore_errors=True); removed += 1

size = sum(os.path.getsize(os.path.join(out, f)) for f in os.listdir(out))
total_rows = sum(v['rows'] for v in manifest['tables'].values())
print(f'백업 완료: 표 {len(manifest["tables"])}개 · {total_rows}건 · {size/1024/1024:.1f}MB')
print(f'위치: {out}')
print(f'문제: {problems if problems else "없음"} · 오래된 폴더 정리 {removed}개')
sys.exit(1 if problems else 0)
