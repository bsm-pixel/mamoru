# 공개 키(anon)만으로 각 표가 읽히는지 전수 확인 — 읽기만 한다. 실행: 앱 폴더에서  PYTHONIOENCODING=utf-8 py scripts/anon-probe.py
import json,sys,urllib.request,urllib.error,re
env={}
for l in open('.env.local',encoding='utf-8'):
    m=re.match(r'^([A-Z0-9_]+)=(.*)$',l.strip())
    if m: env[m.group(1)]=m.group(2).strip().strip('"')
U=env['NEXT_PUBLIC_SUPABASE_URL']; S=env['SUPABASE_SERVICE_ROLE_KEY']; A=env['NEXT_PUBLIC_SUPABASE_ANON_KEY']
def get(path,key,extra=None):
    h={'apikey':key,'Authorization':'Bearer '+key}
    if extra: h.update(extra)
    try:
        r=urllib.request.urlopen(urllib.request.Request(U+path,headers=h),timeout=20)
        return r.status,r.read(),dict(r.headers)
    except urllib.error.HTTPError as e:
        return e.code,e.read(),dict(e.headers)
st,body,_=get('/rest/v1/',S)
spec=json.loads(body)
tables=sorted(k for k in spec.get('definitions',{}).keys())
rpcs=sorted(p[5:] for p in spec.get('paths',{}) if p.startswith('/rpc/'))
readable=[];blocked=[];empty=[]
for t in tables:
    st,b,h=get(f'/rest/v1/{t}?select=*&limit=1',A,{'Prefer':'count=exact'})
    cr=h.get('Content-Range') or h.get('content-range') or ''
    total=cr.split('/')[-1] if '/' in cr else '?'
    if st in (200,206):
        n=len(json.loads(b))
        # 서비스 키로 실제 건수
        st2,b2,h2=get(f'/rest/v1/{t}?select=*&limit=1',S,{'Prefer':'count=exact'})
        cr2=h2.get('Content-Range') or h2.get('content-range') or ''
        real=cr2.split('/')[-1] if '/' in cr2 else '?'
        (readable if (n>0 or (total not in('0','?','*'))) else empty).append((t,total,real))
    else:
        blocked.append((t,st))
print('표',len(tables),'개 / 함수',len(rpcs),'개')
print('\n[공개 키로 읽힘] ',len(readable))
for t,a,r in readable: print(f'  {t}: 공개키 {a}건 / 실제 {r}건')
print('\n[요청은 받지만 0건(행 보안으로 가려짐 또는 빈 표)] ',len(empty))
print('  '+', '.join(f'{t}(실제 {r})' for t,a,r in empty))
print('\n[거부] ',len(blocked)); print('  '+', '.join(f'{t}:{s}' for t,s in blocked))
# 저장소 버킷
st,b,_=get('/storage/v1/bucket',S)
try: print('\n[저장소 버킷]',[(x['name'],'공개' if x.get('public') else '비공개') for x in json.loads(b)])
except Exception as e: print('bucket',st,b[:100])
print('\n[함수 이름]',', '.join(rpcs))
