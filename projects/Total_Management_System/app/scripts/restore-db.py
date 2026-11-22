# TMS 자료 되살리기 — 백업(scripts/backup-db.py)의 표 하나를 운영과 비교하고, 고르면 되살린다 (2026-11-22)
#
# 기본은 「보기만」: 운영 자료를 바꾸지 않는다.
#   PYTHONIOENCODING=utf-8 py scripts/restore-db.py <표이름> [백업폴더]
#     → 백업 파일이 온전한지(_목록.json 의 건수 · 지문과 같은지) 확인
#     → 운영과 비교: 백업에만 있는 줄(지워진 것) · 운영에만 있는 줄(백업 뒤 생긴 것) · 값이 달라진 줄
#     → 되살리면 무엇이 바뀌는지 표로 보여 주고, 되살릴 줄을 <백업폴더>/_되살리기_<표>.json 으로 남긴다
#
# 실제로 되살리기(운영 자료를 바꿈 — 사장님 확인 뒤에만):
#   ... restore-db.py <표이름> [백업폴더] --apply --only=missing|changed|both --yes
#     missing = 백업에만 있는 줄을 다시 넣는다(지워진 것 되살리기, 기존 줄은 안 건드림)
#     changed = 값이 달라진 줄을 백업 값으로 되돌린다
#     백업 뒤 새로 생긴 줄은 어떤 경우에도 지우지 않는다(되살리기로 오늘 자료를 잃지 않게).
#   되살리기 직전에 지금 운영 상태를 <백업폴더>/_되살리기전_<표>_<시각>.json 으로 따로 저장한다(되돌릴 수 있게).
import json, os, re, sys, hashlib, urllib.request, urllib.error, datetime

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
env = {}
for l in open(os.path.join(HERE, '.env.local'), encoding='utf-8'):
    m = re.match(r'^([A-Z0-9_]+)=(.*)$', l.strip())
    if m: env[m.group(1)] = m.group(2).strip().strip('"')
U = env['NEXT_PUBLIC_SUPABASE_URL']; S = env['SUPABASE_SERVICE_ROLE_KEY']
ROOT = r'C:\마모루드라이브\TMS백업'

args = [a for a in sys.argv[1:] if not a.startswith('--')]
flags = {a.split('=')[0]: (a.split('=')[1] if '=' in a else True) for a in sys.argv[1:] if a.startswith('--')}
if not args:
    print('쓰는 법: py scripts/restore-db.py <표이름> [백업폴더] [--apply --only=missing|changed|both --yes]'); sys.exit(1)
table = args[0]
if len(args) > 1:
    folder = args[1]
else:
    ds = sorted(d for d in os.listdir(ROOT) if re.match(r'^\d{4}-\d{2}-\d{2}_\d{4}$', d))
    if not ds: print('백업 폴더가 없습니다'); sys.exit(1)
    folder = os.path.join(ROOT, ds[-1])

def req(path, method='GET', body=None, extra=None):
    h = {'apikey': S, 'Authorization': 'Bearer ' + S, 'Content-Type': 'application/json'}
    if extra: h.update(extra)
    r = urllib.request.urlopen(urllib.request.Request(U + path, data=body, method=method, headers=h), timeout=120)
    return r.read()

# 1) 백업 파일 확인
fp = os.path.join(folder, table + '.json')
if not os.path.exists(fp): print(f'백업에 {table} 표가 없습니다: {fp}'); sys.exit(1)
raw = open(fp, 'rb').read()
backup = json.loads(raw)
man = json.load(open(os.path.join(folder, '_목록.json'), encoding='utf-8')).get('tables', {}).get(table, {})
ok_rows = man.get('rows') == len(backup)
ok_sha = man.get('sha256') == hashlib.sha256(raw).hexdigest()[:16]
print(f'백업: {folder}  표: {table}  {len(backup)}건  (목록과 건수 {"같음" if ok_rows else "다름!"} · 지문 {"같음" if ok_sha else "다름!"})')
if not (ok_rows and ok_sha): print('⚠️ 백업 파일이 목록과 다릅니다 — 되살리지 않습니다'); sys.exit(2)
if not backup: print('백업에 줄이 없습니다 — 비교할 것 없음'); sys.exit(0)
if 'id' not in backup[0]: print('id 칸이 없는 표라 줄 단위로 맞출 수 없습니다 — 되살리지 않습니다'); sys.exit(2)

# 2) 운영 자료 읽기(읽기만)
live, off = [], 0
while True:
    chunk = json.loads(req(f'/rest/v1/{table}?select=*&order=id&limit=1000&offset={off}'))
    live += chunk
    if len(chunk) < 1000: break
    off += 1000
L = {r['id']: r for r in live}; B = {r['id']: r for r in backup}
missing = [B[i] for i in B if i not in L]                 # 백업에만 — 그 뒤 지워진 줄
added = [L[i] for i in L if i not in B]                   # 운영에만 — 백업 뒤 생긴 줄(되살리기가 건드리지 않음)
changed = []
for i in B:
    if i in L:
        diff = [k for k in B[i] if k in L[i] and B[i][k] != L[i][k] and k not in ('updated_at',)]
        if diff: changed.append((B[i], diff))
print(f'운영: {len(live)}건')
print(f'  백업에만 있음(지워진 줄): {len(missing)}건')
print(f'  값이 달라진 줄: {len(changed)}건' + (f'  (자주 바뀐 칸: {", ".join(sorted({k for _, d in changed for k in d}, key=lambda k: -sum(k in d for _, d in changed))[:6])})' if changed else ''))
print(f'  백업 뒤 새로 생긴 줄: {len(added)}건 — 되살리기는 이 줄을 지우지 않습니다')

plan = {'table': table, 'backup': folder, 'missing': missing, 'changed': [r for r, _ in changed]}
out = os.path.join(folder, f'_되살리기_{table}.json')
json.dump(plan, open(out, 'w', encoding='utf-8'), ensure_ascii=False)
print(f'되살릴 줄 목록: {out}')

if not flags.get('--apply'):
    print('(보기만 — 운영 자료는 바꾸지 않았습니다)'); sys.exit(0)

# 3) 실제 되살리기 — 확인 표시가 있어야만
only = flags.get('--only', 'missing')
if flags.get('--yes') is not True:
    print('되살리려면 --yes 를 붙이세요(운영 자료가 바뀝니다)'); sys.exit(3)
rows = (missing if only in ('missing', 'both') else []) + ([r for r, _ in changed] if only in ('changed', 'both') else [])
if not rows: print('되살릴 줄이 없습니다'); sys.exit(0)
stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
json.dump(live, open(os.path.join(folder, f'_되살리기전_{table}_{stamp}.json'), 'w', encoding='utf-8'), ensure_ascii=False)
done = 0
for i in range(0, len(rows), 200):
    req(f'/rest/v1/{table}?on_conflict=id', 'POST', json.dumps(rows[i:i + 200]).encode('utf-8'), {'Prefer': 'resolution=merge-duplicates,return=minimal'})
    done += len(rows[i:i + 200])
print(f'되살림: {done}줄 ({only}). 되살리기 전 상태: _되살리기전_{table}_{stamp}.json')
