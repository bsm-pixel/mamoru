'use client';

import { useCallback, useEffect, useMemo, useState } from 'react';
import Link from 'next/link';
import { ArrowLeft, Star, Trophy, Calendar, ImagePlus, Save, Loader2, ExternalLink } from 'lucide-react';
import { resizeImage } from '@/lib/utils/resize-image';

const LIVE_PAGE = 'https://page.mamoru.kr/projects/reviews/page_review_event.html';

interface Prize { rank: number; name: string; desc: string; image_url: string; count: number }
interface EventConfig {
  deadline: string | null;
  announce_at: string | null;
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

const EMPTY_PRIZES: Prize[] = [
  { rank: 1, name: '', desc: '', image_url: '', count: 1 },
  { rank: 2, name: '', desc: '', image_url: '', count: 0 },
  { rank: 3, name: '', desc: '', image_url: '', count: 0 },
];

export default function ReviewEventPage() {
  const [month, setMonth] = useState<string>(nowYYMM());
  const [config, setConfig] = useState<EventConfig>({ deadline: null, announce_at: null, hero_image_url: null, prizes: EMPTY_PRIZES, status: 'draft' });
  const [reviews, setReviews] = useState<ReviewRow[]>([]);
  const [marks, setMarks] = useState<Record<string, WinnerMark>>({});
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState<string>('');

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
        hero_image_url: cfg.hero_image_url ?? null,
        prizes: (Array.isArray(cfg.prizes) && cfg.prizes.length ? cfg.prizes : EMPTY_PRIZES).map((p: Prize, i: number) => ({ rank: p.rank ?? i + 1, name: p.name ?? '', desc: p.desc ?? '', image_url: p.image_url ?? '', count: p.count ?? 0 })),
        status: cfg.status ?? 'draft',
      } : { deadline: null, announce_at: null, hero_image_url: null, prizes: EMPTY_PRIZES, status: 'draft' });
      const rv: ReviewRow[] = data.reviews || [];
      setReviews(rv);
      const mk: Record<string, WinnerMark> = {};
      rv.forEach((r) => {
        if (r.event_month === m && r.event_rank) {
          mk[r.id] = { rank: r.event_rank, display_name: r.event_display_name || '', route: r.event_route || '' };
        }
      });
      setMarks(mk);
    } catch (e) {
      setMsg('불러오기 실패: ' + String(e));
    } finally { setLoading(false); }
  }, []);

  useEffect(() => { load(month); }, [month, load]);

  const counts = useMemo(() => {
    const c = { 1: 0, 2: 0, 3: 0 } as Record<number, number>;
    Object.values(marks).forEach((m) => { c[m.rank] = (c[m.rank] || 0) + 1; });
    return c;
  }, [marks]);

  function setRank(id: string, rank: number | null) {
    setMarks((prev) => {
      const next = { ...prev };
      if (rank === null) delete next[id];
      else next[id] = { rank, display_name: next[id]?.display_name || '', route: next[id]?.route || '' };
      return next;
    });
  }
  function setMarkField(id: string, field: 'display_name' | 'route', value: string) {
    setMarks((prev) => prev[id] ? { ...prev, [id]: { ...prev[id], [field]: value } } : prev);
  }
  function setPrize(idx: number, field: keyof Prize, value: string | number) {
    setConfig((c) => ({ ...c, prizes: c.prizes.map((p, i) => i === idx ? { ...p, [field]: value } : p) }));
  }

  async function uploadImage(file: File): Promise<string | null> {
    try {
      const resized = await resizeImage(file, 1200, 0.85);
      const fd = new FormData();
      fd.append('files', resized);
      const res = await fetch('/api/reviews/upload-bulk', { method: 'POST', body: fd });
      const data = await res.json();
      const url = data?.results?.[0]?.url;
      return url || null;
    } catch { return null; }
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

          <label className="block text-xs text-stone-500 mb-1">응모 마감일 (히어로 카운트다운)</label>
          <input type="datetime-local" value={toLocalInput(config.deadline)} onChange={(e) => setConfig((c) => ({ ...c, deadline: fromLocal(e.target.value) }))} className="w-full mb-3 px-3 py-2 rounded-lg border border-stone-200 text-sm" />

          <label className="block text-xs text-stone-500 mb-1">발표 예정일 (안내용, 선택)</label>
          <input type="datetime-local" value={toLocalInput(config.announce_at)} onChange={(e) => setConfig((c) => ({ ...c, announce_at: fromLocal(e.target.value) }))} className="w-full mb-3 px-3 py-2 rounded-lg border border-stone-200 text-sm" />

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
                <div className="flex items-center gap-2">
                  {p.image_url
                    ? <img src={p.image_url} alt="" className="w-12 h-12 rounded object-cover border border-stone-200" />
                    : <div className="w-12 h-12 rounded border border-dashed border-stone-300 flex items-center justify-center text-stone-300"><ImagePlus size={16} /></div>}
                  <label className="text-xs text-stone-600 border border-stone-200 rounded px-2.5 py-1.5 cursor-pointer hover:bg-white">
                    이미지 업로드
                    <input type="file" accept="image/*" className="hidden" onChange={async (e) => { const f = e.target.files?.[0]; if (f) { const url = await uploadImage(f); if (url) setPrize(idx, 'image_url', url); } }} />
                  </label>
                </div>
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
      <div className="mt-5 flex items-center gap-2 flex-wrap sticky bottom-0 bg-white/90 backdrop-blur py-3 border-t border-stone-100">
        <span className="text-xs text-stone-500 mr-1">현재 상태:
          <span className={`ml-1 font-semibold ${config.status === 'announced' ? 'text-emerald-600' : config.status === 'live' ? 'text-blue-600' : 'text-stone-400'}`}>
            {config.status === 'announced' ? '발표됨(공개)' : config.status === 'live' ? '진행중(공개)' : '임시저장(비공개)'}
          </span>
        </span>
        <div className="ml-auto flex items-center gap-2">
          <button disabled={saving} onClick={() => save('draft')} className="px-3 py-2 rounded-lg border border-stone-200 text-sm hover:bg-stone-50 disabled:opacity-50">임시저장</button>
          <button disabled={saving} onClick={() => save('live')} className="px-3 py-2 rounded-lg bg-blue-600 text-white text-sm hover:bg-blue-700 disabled:opacity-50 flex items-center gap-1"><Save size={14} />진행중으로 게시</button>
          <button disabled={saving} onClick={() => save('announced')} className="px-3 py-2 rounded-lg bg-emerald-600 text-white text-sm hover:bg-emerald-700 disabled:opacity-50 flex items-center gap-1">{saving ? <Loader2 size={14} className="animate-spin" /> : <Trophy size={14} />}당첨 발표</button>
        </div>
      </div>
      <p className="text-[11px] text-stone-400 mt-2 leading-relaxed">
        · <b>진행중으로 게시</b> = 이달의 상품·마감일이 고객 페이지 Hero/이달의 상품에 노출됩니다.<br />
        · <b>당첨 발표</b> = 선정한 당첨자가 고객 페이지 &apos;지난 당첨자&apos; 아카이브(해당 월 탭)에 마스킹되어 노출됩니다.<br />
        · 표시명을 비우면 이름이 자동 마스킹(홍**님), 전화 뒷자리도 자동 마스킹(010-****-32**)됩니다.
      </p>
    </div>
  );
}
