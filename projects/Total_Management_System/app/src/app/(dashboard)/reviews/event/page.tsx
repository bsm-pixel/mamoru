'use client';

import { useCallback, useEffect, useMemo, useState } from 'react';
import Link from 'next/link';
import { ArrowLeft, Star, Trophy, Calendar, ImagePlus, Save, Loader2, ExternalLink } from 'lucide-react';
import { resizeImage } from '@/lib/utils/resize-image';

const LIVE_PAGE = 'https://page.mamoru.kr/projects/reviews/page_review_event.html';

interface Prize { rank: number; name: string; desc: string; image_url: string; count: number; image_urls?: string[] }
interface EventConfig {
  deadline: string | null;
  announce_at: string | null;
  entry_start: string | null;   // 응모 시작일(선택). null=그 달 1일부터
  hero_image_url: string | null;
  prizes: Prize[];
  status: 'draft' | 'live' | 'announced';
}
interface ReviewRow {
  id: string; review_id: string | null; type: string; subtype: string | null;
  name: string | null; stars: number | null; content: string | null;
  photo_urls: string[] | null; product: string | null; created_at: string; status: string;
  event_month: string | null; event_rank: number | null;
  event_display_name: string | null; event_route: string | null;
}
interface WinnerMark { rank: number; display_name: string; route: string }

function nowYYMM(): string {
  const d = new Date();
  return String(d.getFullYear()).slice(2) + String(d.getMonth() + 1).padStart(2, '0');
}
function shiftMonth(yymm: string, delta: number): string {
  const y = 2000 + parseInt(yymm.slice(0, 2), 10);
  const m = parseInt(yymm.slice(2, 4), 10) - 1 + delta;
  const d = new Date(y, m, 1);
  return String(d.getFullYear()).slice(2) + String(d.getMonth() + 1).padStart(2, '0');
}
function monthTitle(yymm: string): string {
  return `20${yymm.slice(0, 2)}년 ${parseInt(yymm.slice(2, 4), 10)}월`;
}
/** timestamptz(UTC ISO) → datetime-local 값(KST 벽시계) */
function toLocalInput(iso: string | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  const kst = new Date(d.getTime() + 9 * 3600 * 1000);
  return kst.toISOString().slice(0, 16);
}
/** datetime-local 문자열을 KST 로 해석해 UTC ISO 반환 */
function fromLocal(v: string): string | null {
  if (!v) return null;
  // v = 'YYYY-MM-DDTHH:mm' (KST 벽시계) → UTC = KST - 9h
  const [date, time] = v.split('T');
  const [Y, Mo, D] = date.split('-').map(Number);
  const [H, Mi] = time.split(':').map(Number);
  const utcMs = Date.UTC(Y, Mo - 1, D, H, Mi) - 9 * 3600 * 1000;
  return new Date(utcMs).toISOString();
}
/** UTC ISO → 'YYYY-MM-DD'(KST 날짜) — <input type="date"> 값 */
function isoToKstDate(iso: string | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  return new Date(d.getTime() + 9 * 3600 * 1000).toISOString().slice(0, 10);
}
/** 'YYYY-MM-DD'(KST 날짜) → 그 날 00:00 KST 의 UTC ISO */
function kstDateToISO(dateStr: string): string | null {
  if (!dateStr) return null;
  const [Y, M, D] = dateStr.split('-').map(Number);
  return new Date(Date.UTC(Y, M - 1, D) - 9 * 3600 * 1000).toISOString();
}

const EMPTY_PRIZES: Prize[] = [
  { rank: 1, name: '', desc: '', image_url: '', count: 1, image_urls: [] },
  { rank: 2, name: '', desc: '', image_url: '', count: 0, image_urls: [] },
  { rank: 3, name: '', desc: '', image_url: '', count: 0, image_urls: [] },
];
// 상품 이미지 배열 — image_urls(신) 우선, 없으면 image_url(구) 폴백(하위호환)
function prizeImgList(p: Prize): string[] { return (Array.isArray(p.image_urls) && p.image_urls.length) ? p.image_urls.filter(Boolean) : (p.image_url ? [p.image_url] : []); }

export default function ReviewEventPage() {
  const [month, setMonth] = useState<string>(nowYYMM());
  const [config, setConfig] = useState<EventConfig>({ deadline: null, announce_at: null, entry_start: null, hero_image_url: null, prizes: EMPTY_PRIZES, status: 'draft' });
  const [reviews, setReviews] = useState<ReviewRow[]>([]);
  const [marks, setMarks] = useState<Record<string, WinnerMark>>({});
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState<string>('');
  const [dirty, setDirty] = useState(false);   // 저장 안 된 변경사항 여부

  const load = useCallback(async (m: string) => {
    setLoading(true); setMsg('');
    try {
      const res = await fetch(`/api/reviews/event?month=${m}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || '조회 실패');
      const cfg = data.config;
      setConfig(cfg ? {
        deadline: cfg.deadline ?? null,
        announce_at: cfg.announce_at ?? null,
        entry_start: cfg.entry_start ?? null,
        hero_image_url: cfg.hero_image_url ?? null,
        prizes: (Array.isArray(cfg.prizes) && cfg.prizes.length ? cfg.prizes : EMPTY_PRIZES).map((p: Prize, i: number) => { const imgs = prizeImgList(p); return { rank: p.rank ?? i + 1, name: p.name ?? '', desc: p.desc ?? '', image_url: imgs[0] ?? '', count: p.count ?? 0, image_urls: imgs }; }),
        status: cfg.status ?? 'draft',
      } : { deadline: null, announce_at: null, entry_start: null, hero_image_url: null, prizes: EMPTY_PRIZES, status: 'draft' });
      const rv: ReviewRow[] = data.reviews || [];
      setReviews(rv);
      const mk: Record<string, WinnerMark> = {};
      rv.forEach((r) => {
        if (r.event_month === m && r.event_rank) {
          mk[r.id] = { rank: r.event_rank, display_name: r.event_display_name || '', route: r.event_route || '' };
        }
      });
      setMarks(mk);
      setDirty(false);   // 새로 불러오면 깨끗한 상태
    } catch (e) {
      setMsg('불러오기 실패: ' + String(e));
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { load(month); }, [month, load]);

  // 저장 안 한 채 새로고침/닫기 시 경고
  useEffect(() => {
    if (!dirty) return;
    const h = (e: BeforeUnloadEvent) => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', h);
    return () => window.removeEventListener('beforeunload', h);
  }, [dirty]);

  // 응모 시작일 변경 시 응모자 풀만 다시 불러옴(설정은 유지). 사용자가 만지던 마킹은 풀에 남은 것만 보존
  const reloadPool = useCallback(async (m: string, startDate: string) => {
    setLoading(true);
    try {
      const q = startDate ? `&start=${startDate}` : '';
      const res = await fetch(`/api/reviews/event?month=${m}${q}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || '조회 실패');
      const rv: ReviewRow[] = data.reviews || [];
      setReviews(rv);
      setMarks((prev) => {
        const next: Record<string, WinnerMark> = {};
        rv.forEach((r) => {
          if (prev[r.id]) next[r.id] = prev[r.id];                                   // 사용자가 만지던 마킹 보존
          else if (r.event_month === m && r.event_rank) next[r.id] = { rank: r.event_rank, display_name: r.event_display_name || '', route: r.event_route || '' };
        });
        return next;
      });
    } catch (e) {
      setMsg('불러오기 실패: ' + String(e));
    } finally { setLoading(false); }
  }, []);

  const counts = useMemo(() => {
    const c = { 1: 0, 2: 0, 3: 0 } as Record<number, number>;
    Object.values(marks).forEach((m) => { c[m.rank] = (c[m.rank] || 0) + 1; });
    return c;
  }, [marks]);

  function setRank(id: string, rank: number | null) {
    setDirty(true);
    setMarks((prev) => {
      const next = { ...prev };
      if (rank === null) delete next[id];
      else next[id] = { rank, display_name: next[id]?.display_name || '', route: next[id]?.route || '' };
      return next;
    });
  }
  function setMarkField(id: string, field: 'display_name' | 'route', value: string) {
    setDirty(true);
    setMarks((prev) => prev[id] ? { ...prev, [id]: { ...prev[id], [field]: value } } : prev);
  }
  function setPrize(idx: number, field: keyof Prize, value: string | number) {
    setDirty(true);
    setConfig((c) => ({ ...c, prizes: c.prizes.map((p, i) => i === idx ? { ...p, [field]: value } : p) }));
  }
  function setPrizeImages(idx: number, urls: string[]) {
    setDirty(true);
    setConfig((c) => ({ ...c, prizes: c.prizes.map((p, i) => i === idx ? { ...p, image_urls: urls, image_url: urls[0] || '' } : p) }));
  }
  function removePrizeImage(idx: number, imgIdx: number) {
    setDirty(true);
    setConfig((c) => ({ ...c, prizes: c.prizes.map((p, i) => { if (i !== idx) return p; const arr = prizeImgList(p).slice(); arr.splice(imgIdx, 1); return { ...p, image_urls: arr, image_url: arr[0] || '' }; }) }));
  }
  // 게시 상태 변경(+저장) — 공개/발표는 확인
  function changeStatus(s: EventConfig['status']) {
    if (s === config.status) { save(s); return; }
    if (s === 'announced' && !window.confirm('당첨자를 고객 페이지에 공개할까요?')) return;
    if (s === 'live' && !window.confirm('이달의 이벤트를 고객에게 공개(진행중)할까요?')) return;
    if (s === 'draft' && config.status !== 'draft' && !window.confirm('고객 페이지에서 내리고 비공개로 전환할까요?')) return;
    save(s);
  }

  async function uploadImages(files: File[]): Promise<string[]> {
    try {
      const fd = new FormData();
      for (const f of files) { const r = await resizeImage(f, 1200, 0.85); fd.append('files', r); }
      const res = await fetch('/api/reviews/upload-bulk', { method: 'POST', body: fd });
      const data = await res.json();
      return (Array.isArray(data?.results) ? data.results : []).map((x: { url?: string }) => x.url).filter(Boolean) as string[];
    } catch { return []; }
  }

  async function save(nextStatus?: EventConfig['status']) {
    setSaving(true); setMsg('');
    try {
      const status = nextStatus || config.status;
      const winners = Object.entries(marks).map(([id, m]) => ({ id, rank: m.rank, display_name: m.display_name, route: m.route }));
      const res = await fetch('/api/reviews/event', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ month, config: { ...config, status }, winners }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || '저장 실패');
      setConfig((c) => ({ ...c, status }));
      setDirty(false);
      setMsg(`저장 완료 (${status === 'announced' ? '발표됨 — 고객 페이지 지난 당첨자에 노출' : status === 'live' ? '진행중 — 고객 페이지 이달의 상품에 노출' : '임시저장(비공개)'})`);
    } catch (e) {
      setMsg('저장 실패: ' + String(e));
    } finally { setSaving(false); }
  }

  return (
    <div className="p-4 md:p-6 max-w-6xl mx-auto">
      {/* 헤더 */}
      <div className="flex items-center gap-3 mb-4 flex-wrap">
        <Link href="/reviews" className="text-stone-500 hover:text-stone-800 flex items-center gap-1 text-sm"><ArrowLeft size={16} />리뷰관리</Link>
        <h1 className="text-xl font-bold text-stone-900 flex items-center gap-2"><Trophy size={20} className="text-amber-500" />리뷰 이벤트 관리</h1>
        <a href={LIVE_PAGE} target="_blank" rel="noreferrer" className="ml-auto text-xs text-stone-500 hover:text-stone-800 flex items-center gap-1"><ExternalLink size={13} />고객 페이지</a>
      </div>

      {/* 월 선택 */}
      <div className="flex items-center gap-2 mb-4">
        <button onClick={() => setMonth(shiftMonth(month, -1))} className="px-3 py-1.5 rounded-lg border border-stone-200 text-sm hover:bg-stone-50">← 이전달</button>
        <div className="px-4 py-1.5 rounded-lg bg-stone-900 text-white text-sm font-semibold min-w-[120px] text-center">{monthTitle(month)}</div>
        <button onClick={() => setMonth(shiftMonth(month, 1))} className="px-3 py-1.5 rounded-lg border border-stone-200 text-sm hover:bg-stone-50">다음달 →</button>
        {loading && <Loader2 size={16} className="animate-spin text-stone-400" />}
      </div>

      {msg && <div className={`mb-4 px-4 py-2 rounded-lg text-sm ${msg.includes('실패') ? 'bg-rose-100 text-rose-700 font-medium' : 'bg-stone-100 text-stone-700'}`}>{msg}</div>}

      <div className="grid lg:grid-cols-2 gap-5">
        {/* ── 이벤트 설정 ── */}
        <section className="border border-stone-200 rounded-xl p-4 bg-white">
          <h2 className="font-semibold text-stone-800 mb-3 flex items-center gap-1.5"><Calendar size={16} />이벤트 설정</h2>

          <label className="block text-xs text-stone-500 mb-1">응모 시작일 (선택 · 응모자 집계 하한)</label>
          <div className="flex items-center gap-2 mb-1">
            <input type="date" value={isoToKstDate(config.entry_start)} onChange={(e) => { const iso = kstDateToISO(e.target.value); setDirty(true); setConfig((c) => ({ ...c, entry_start: iso })); reloadPool(month, e.target.value); }} className="flex-1 px-3 py-2 rounded-lg border border-stone-200 text-sm" />
            {config.entry_start && <button type="button" onClick={() => { setDirty(true); setConfig((c) => ({ ...c, entry_start: null })); reloadPool(month, ''); }} className="px-2.5 py-2 rounded-lg border border-stone-200 text-xs text-stone-500 hover:bg-stone-50">지우기</button>}
          </div>
          <p className="text-[11px] text-stone-400 mb-3">비우면 <b>{monthTitle(month)} 1일</b>부터 집계. 첫 회차처럼 과거 후기까지 포함하려면 시작일을 앞당겨 지정하세요. (끝은 항상 그 달 말일)</p>

          <label className="block text-xs text-stone-500 mb-1">응모 마감일 (히어로 카운트다운)</label>
          <input type="datetime-local" value={toLocalInput(config.deadline)} onChange={(e) => { setDirty(true); setConfig((c) => ({ ...c, deadline: fromLocal(e.target.value) })); }} className="w-full mb-3 px-3 py-2 rounded-lg border border-stone-200 text-sm" />

          <label className="block text-xs text-stone-500 mb-1">발표 예정일 (안내용, 선택)</label>
          <input type="datetime-local" value={toLocalInput(config.announce_at)} onChange={(e) => { setDirty(true); setConfig((c) => ({ ...c, announce_at: fromLocal(e.target.value) })); }} className="w-full mb-3 px-3 py-2 rounded-lg border border-stone-200 text-sm" />

          {/* 상품 1/2/3등 */}
          <div className="space-y-3 mt-2">
            {config.prizes.map((p, idx) => (
              <div key={p.rank} className="border border-stone-100 rounded-lg p-3 bg-stone-50/60">
                <div className="flex items-center gap-2 mb-2">
                  <span className="text-xs font-bold text-white bg-stone-800 rounded-full px-2 py-0.5">{p.rank}등</span>
                  <input type="number" min={0} value={p.count} onChange={(e) => setPrize(idx, 'count', parseInt(e.target.value) || 0)} className="w-16 px-2 py-1 rounded border border-stone-200 text-sm" />
                  <span className="text-xs text-stone-500">명 추첨</span>
                </div>
                <input placeholder="상품명 (예: 프리미엄 가위 1자루)" value={p.name} onChange={(e) => setPrize(idx, 'name', e.target.value)} className="w-full mb-2 px-2.5 py-1.5 rounded border border-stone-200 text-sm" />
                <input placeholder="한 줄 설명 (선택)" value={p.desc} onChange={(e) => setPrize(idx, 'desc', e.target.value)} className="w-full mb-2 px-2.5 py-1.5 rounded border border-stone-200 text-sm" />
                <div className="flex items-center gap-2 flex-wrap">
                  {prizeImgList(p).map((u, ui) => (
                    <div key={ui} className="relative">
                      <img src={u} alt="" className="w-12 h-12 rounded object-cover border border-stone-200" />
                      <button type="button" onClick={() => removePrizeImage(idx, ui)} aria-label="삭제"
                        className="absolute -top-1.5 -right-1.5 w-4 h-4 rounded-full bg-stone-800 text-white text-[10px] leading-none flex items-center justify-center">×</button>
                    </div>
                  ))}
                  {prizeImgList(p).length === 0 && <div className="w-12 h-12 rounded border border-dashed border-stone-300 flex items-center justify-center text-stone-300"><ImagePlus size={16} /></div>}
                  <label className="text-xs text-stone-600 border border-stone-200 rounded px-2.5 py-1.5 cursor-pointer hover:bg-white">
                    이미지 추가
                    <input type="file" accept="image/*" multiple className="hidden" onChange={async (e) => { const fs = Array.from(e.target.files || []); if (fs.length) { const urls = await uploadImages(fs); if (urls.length) setPrizeImages(idx, prizeImgList(p).concat(urls)); } }} />
                  </label>
                </div>
                {prizeImgList(p).length > 1 && <p className="text-[11px] text-stone-400 mt-1">여러 장이면 고객 화면에서 좌우로 넘겨봅니다 · 첫 장이 대표 이미지</p>}
              </div>
            ))}
          </div>
        </section>

        {/* ── 응모자(그 달 후기) 선정 ── */}
        <section className="border border-stone-200 rounded-xl p-4 bg-white">
          <div className="flex items-center justify-between mb-3">
            <h2 className="font-semibold text-stone-800 flex items-center gap-1.5"><Star size={16} />응모자 · 등수 선정</h2>
            <span className="text-xs text-stone-500">1등 {counts[1] || 0} · 2등 {counts[2] || 0} · 3등 {counts[3] || 0}</span>
          </div>

          {reviews.length === 0 && !loading && <div className="text-sm text-stone-400 py-8 text-center">이 달에 등록된 후기가 없습니다.</div>}

          <div className="space-y-3 max-h-[560px] overflow-y-auto pr-1">
            {reviews.map((r) => {
              const mk = marks[r.id];
              return (
                <div key={r.id} className={`border rounded-lg p-3 ${mk ? 'border-amber-300 bg-amber-50/40' : 'border-stone-100'}`}>
                  <div className="flex items-center gap-2 mb-1">
                    <span className="font-semibold text-sm text-stone-800">{r.name || '이름없음'}</span>
                    <span className="text-amber-500 text-xs">{'★'.repeat(r.stars || 0)}<span className="text-stone-300">{'★'.repeat(5 - (r.stars || 0))}</span></span>
                    {r.product && <span className="text-xs text-stone-400">· {r.product}</span>}
                    {r.status !== 'approved' && <span className="text-[10px] text-rose-500 border border-rose-200 rounded px-1">{r.status}</span>}
                  </div>
                  <p className="text-xs text-stone-600 line-clamp-2 mb-2">{r.content}</p>
                  {r.photo_urls && r.photo_urls.length > 0 && (
                    <div className="flex gap-1 mb-2">{r.photo_urls.slice(0, 4).map((u, i) => <img key={i} src={u} alt="" className="w-9 h-9 rounded object-cover border border-stone-200" />)}</div>
                  )}
                  {/* 등수 토글 */}
                  <div className="flex items-center gap-1">
                    {[1, 2, 3].map((rank) => (
                      <button key={rank} onClick={() => setRank(r.id, mk?.rank === rank ? null : rank)}
                        className={`px-2.5 py-1 rounded text-xs font-medium ${mk?.rank === rank ? 'bg-stone-900 text-white' : 'bg-stone-100 text-stone-500 hover:bg-stone-200'}`}>{rank}등</button>
                    ))}
                    {mk && <button onClick={() => setRank(r.id, null)} className="px-2 py-1 rounded text-xs text-rose-500 hover:bg-rose-50">제외</button>}
                  </div>
                  {mk && (
                    <div className="grid grid-cols-2 gap-2 mt-2">
                      <input placeholder="표시명 (비우면 홍**님 자동)" value={mk.display_name} onChange={(e) => setMarkField(r.id, 'display_name', e.target.value)} className="px-2 py-1 rounded border border-stone-200 text-xs" />
                      <input placeholder="경로 (예: 매장 구매 · 블런트)" value={mk.route} onChange={(e) => setMarkField(r.id, 'route', e.target.value)} className="px-2 py-1 rounded border border-stone-200 text-xs" />
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </section>
      </div>

      {/* 저장 / 상태 */}
      <div className="mt-5 flex items-center gap-3 flex-wrap sticky bottom-0 bg-white/90 backdrop-blur py-3 border-t border-stone-100">
        {/* 저장 상태 표시 */}
        {dirty
          ? <span className="text-xs font-semibold text-amber-600 flex items-center gap-1.5"><span className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse" />저장 안 됨</span>
          : <span className="text-xs text-stone-400 flex items-center gap-1.5"><span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />저장됨</span>}

        {/* 게시 상태 세그먼트 (클릭 시 상태 전환 + 저장) */}
        <div className="flex items-center gap-1.5">
          <span className="text-xs text-stone-500">게시</span>
          <div className="inline-flex rounded-lg border border-stone-200 overflow-hidden">
            {([['draft', '비공개'], ['live', '진행중'], ['announced', '발표됨']] as const).map(([s, label], i) => {
              const active = config.status === s;
              const activeCls = s === 'draft' ? 'bg-stone-700 text-white' : s === 'live' ? 'bg-blue-600 text-white' : 'bg-emerald-600 text-white';
              return (
                <button key={s} disabled={saving} onClick={() => changeStatus(s)}
                  className={`px-3 py-1.5 text-xs font-semibold transition disabled:opacity-50 ${i > 0 ? 'border-l border-stone-200' : ''} ${active ? activeCls : 'bg-white text-stone-500 hover:bg-stone-50'}`}>
                  {label}
                </button>
              );
            })}
          </div>
        </div>

        {/* 저장 (현재 상태 유지) — 평소엔 이것만 */}
        <button disabled={saving} onClick={() => save()}
          className="ml-auto px-4 py-2 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 disabled:opacity-50 flex items-center gap-1.5">
          {saving ? <Loader2 size={15} className="animate-spin" /> : <Save size={15} />}저장
        </button>
      </div>
      <p className="text-[11px] text-stone-400 mt-2 leading-relaxed">
        · <b>저장</b> = 편집한 내용을 <b>현재 게시 상태 그대로</b> 저장합니다. (이미지·상품·당첨자 변경은 저장해야 반영돼요)<br />
        · <b>게시</b> 전환 — <b>비공개</b>(고객에 안 보임) → <b>진행중</b>(이달의 상품·마감일 공개, 응모 시작) → <b>발표됨</b>(당첨자 공개).<br />
        · 표시명을 비우면 이름이 자동 마스킹(홍**님), 전화 뒷자리도 자동 마스킹(010-****-32**)됩니다.
      </p>
    </div>
  );
}
