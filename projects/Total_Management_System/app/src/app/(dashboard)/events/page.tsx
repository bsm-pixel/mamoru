'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useEvents, useEventPatch, useEventDelete, useCampaigns, useCreateCampaign, useUpdateCampaign } from '@/hooks/use-events';
import { EVENT_STATUS_LABEL, CAMPAIGN_TYPE_LABEL, type EventSubmission, type EventStatus, type EventCampaign, type DiscountRule } from '@/lib/event/types';
import { Zap, Loader2, Package, Truck, Store, ArrowLeft, Plus, ExternalLink, Settings, X } from 'lucide-react';

const TABS: { key: EventStatus; label: string }[] = [
  { key: 'received', label: '신규접수' },
  { key: 'payment_noticed', label: '입금대기' },
  { key: 'converted', label: '판매전환' },
  { key: 'cancelled', label: '취소' },
];

const won = (n: number) => `${(n || 0).toLocaleString()}원`;
const fmtPhone = (p: string | null) => (p || '').replace(/(\d{3})(\d{3,4})(\d{4})/, '$1-$2-$3');

export default function EventsPage() {
  const router = useRouter();
  const [campaignId, setCampaignId] = useState<string | null>(null); // null = 캠페인 카드 화면
  const [tab, setTab] = useState<EventStatus>('received');
  const [selId, setSelId] = useState<string | null>(null);
  const [showNew, setShowNew] = useState(false);
  const [editCampaign, setEditCampaign] = useState<EventCampaign | null>(null);
  const { data: campaigns, isLoading: campLoading } = useCampaigns();
  const { data: all, isLoading } = useEvents('all');
  const patch = useEventPatch();
  const del = useEventDelete();
  const createCampaign = useCreateCampaign();
  const updateCampaign = useUpdateCampaign();

  // 캠페인별 상태 카운트
  const countsByCampaign = useMemo(() => {
    const m: Record<string, Record<string, number>> = {};
    (all || []).forEach((e) => {
      const cid = e.campaign_id || '_none';
      if (!m[cid]) m[cid] = {};
      m[cid][e.status] = (m[cid][e.status] || 0) + 1;
    });
    return m;
  }, [all]);

  const activeCampaign = useMemo(() => (campaigns || []).find((c) => c.id === campaignId) || null, [campaigns, campaignId]);
  const list = useMemo(() => (all || []).filter((e) => e.campaign_id === campaignId && e.status === tab), [all, campaignId, tab]);
  const tabCounts = countsByCampaign[campaignId || '_none'] || {};
  const sel = useMemo(() => (all || []).find((e) => e.id === selId) || null, [all, selId]);

  // ── 캠페인 카드 화면 ──
  if (!campaignId) {
    return (
      <>
        <Topbar title="EVENT" />
        <div className="min-h-screen bg-neutral-50 px-4 md:px-6 py-4 space-y-4 overflow-x-hidden">
          <div className="flex items-start gap-2 rounded-xl bg-indigo-50 border border-indigo-100 px-3 py-2.5 text-xs text-indigo-700">
            <Zap size={14} className="shrink-0 mt-0.5" />
            <p>진행 중인 이벤트(캠페인)별로 접수를 관리합니다. 매장 방문 즉시구매는 <button onClick={() => router.push('/sales/new')} className="underline font-semibold">판매입력</button>에서 EVENT 품목을 선택하세요.</p>
          </div>

          <div className="flex items-center justify-between">
            <h2 className="text-sm font-bold text-neutral-900">캠페인</h2>
            <button onClick={() => setShowNew(true)} className="flex items-center gap-1 text-xs font-semibold text-white bg-neutral-900 px-3 py-1.5 rounded-lg">
              <Plus size={14} />새 캠페인
            </button>
          </div>

          {campLoading ? (
            <div className="py-16 text-center text-neutral-400"><Loader2 size={20} className="animate-spin inline" /></div>
          ) : (campaigns || []).length === 0 ? (
            <div className="py-16 text-center text-sm text-neutral-400">캠페인이 없습니다. ‘새 캠페인’으로 만들어 주세요.</div>
          ) : (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {(campaigns || []).map((c) => {
                const cc = countsByCampaign[c.id] || {};
                const formUrl = `https://page.mamoru.kr/projects/event/page_form.html?campaign=${c.id}`;
                return (
                  <div key={c.id} onClick={() => { setCampaignId(c.id); setTab('received'); }}
                    className="cursor-pointer text-left bg-white rounded-2xl border border-neutral-200 p-4 hover:border-neutral-400 transition">
                    <div className="flex items-center justify-between gap-2">
                      <div className="min-w-0">
                        <div className="text-[11px] text-neutral-400">{CAMPAIGN_TYPE_LABEL[c.type] || c.type}</div>
                        <div className="text-base font-bold text-neutral-900 truncate">{c.name}</div>
                      </div>
                      <div className="shrink-0 flex items-center gap-1.5">
                        <span className={`text-[10px] px-2 py-0.5 rounded-full ${c.status === 'active' ? 'bg-emerald-50 text-emerald-700' : 'bg-neutral-100 text-neutral-400'}`}>
                          {c.status === 'active' ? '진행중' : '종료'}
                        </span>
                        <button onClick={(e) => { e.stopPropagation(); setEditCampaign(c); }} title="할인·설정"
                          className="w-7 h-7 flex items-center justify-center rounded-lg text-neutral-400 hover:bg-neutral-100"><Settings size={15} /></button>
                      </div>
                    </div>
                    {(c.discount_rules || []).length > 0 && (
                      <div className="mt-2 text-[11px] text-indigo-600">묶음 할인 {c.discount_rules.length}건 설정됨</div>
                    )}
                    <div className="mt-3 grid grid-cols-4 gap-1 text-center">
                      {TABS.map((t) => (
                        <div key={t.key} className="rounded-lg bg-neutral-50 py-2">
                          <div className="text-lg font-bold text-neutral-900">{cc[t.key] || 0}</div>
                          <div className="text-[10px] text-neutral-400">{t.label}</div>
                        </div>
                      ))}
                    </div>
                    {/* 접수페이지 바로가기 — 고객이 보는 폼을 새 탭으로 열어 빠르게 점검 */}
                    <div className="mt-3 flex justify-end">
                      <a href={formUrl} target="_blank" rel="noreferrer" onClick={(e) => e.stopPropagation()}
                        className="inline-flex items-center gap-1 text-[11px] font-semibold text-indigo-600 bg-indigo-50 hover:bg-indigo-100 px-2.5 py-1 rounded-full transition">
                        <ExternalLink size={11} />접수 페이지 열기
                      </a>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {showNew && <CampaignFormModal onClose={() => setShowNew(false)} create={createCampaign} update={updateCampaign} />}
        {editCampaign && <CampaignFormModal campaign={editCampaign} onClose={() => setEditCampaign(null)} create={createCampaign} update={updateCampaign} />}
      </>
    );
  }

  // ── 캠페인 상세(접수 목록) 화면 ──
  return (
    <>
      <Topbar title="EVENT" />
      <div className="min-h-screen bg-neutral-50 px-4 md:px-6 py-4 space-y-4 overflow-x-hidden">
        <button onClick={() => { setCampaignId(null); setSelId(null); }} className="flex items-center gap-1 text-sm text-neutral-500 hover:text-neutral-900">
          <ArrowLeft size={16} />캠페인 목록
        </button>
        <div className="flex items-center justify-between gap-2 flex-wrap">
          <h2 className="text-lg font-bold text-neutral-900">{activeCampaign?.name || '캠페인'}</h2>
          <div className="flex items-center gap-2">
            {activeCampaign && (
              <button onClick={() => setEditCampaign(activeCampaign)}
                className="inline-flex items-center gap-1 text-[11px] font-semibold text-neutral-600 bg-neutral-100 hover:bg-neutral-200 px-2.5 py-1 rounded-full transition">
                <Settings size={11} />할인·설정
              </button>
            )}
            {campaignId && (
              <a href={`https://page.mamoru.kr/projects/event/page_form.html?campaign=${campaignId}`} target="_blank" rel="noreferrer"
                className="inline-flex items-center gap-1 text-[11px] font-semibold text-indigo-600 bg-indigo-50 hover:bg-indigo-100 px-2.5 py-1 rounded-full transition">
                <ExternalLink size={11} />접수 페이지 열기
              </a>
            )}
          </div>
        </div>

        {/* 탭 */}
        <div className="flex gap-1 overflow-x-auto scrollbar-hide border-b border-neutral-200">
          {TABS.map((t) => (
            <button key={t.key} onClick={() => setTab(t.key)}
              className={`shrink-0 whitespace-nowrap px-3 py-2.5 text-sm font-semibold border-b-2 transition ${
                tab === t.key ? 'border-neutral-900 text-neutral-900' : 'border-transparent text-neutral-400 hover:text-neutral-600'
              }`}>
              {t.label}
              {(tabCounts[t.key] || 0) > 0 && <span className="ml-1.5 text-xs px-1.5 py-0.5 rounded-full bg-neutral-200 text-neutral-600">{tabCounts[t.key]}</span>}
            </button>
          ))}
        </div>

        {isLoading ? (
          <div className="py-16 text-center text-sm text-neutral-400"><Loader2 size={20} className="animate-spin inline" /></div>
        ) : list.length === 0 ? (
          <div className="py-16 text-center text-sm text-neutral-400">{EVENT_STATUS_LABEL[tab]} 건이 없습니다</div>
        ) : (
          <div className="space-y-2">
            {list.map((e) => (
              <button key={e.id} onClick={() => setSelId(e.id)}
                className="w-full text-left bg-white rounded-xl border border-neutral-200 px-4 py-3 hover:border-neutral-400 transition">
                <div className="flex items-center justify-between gap-2">
                  <div className="min-w-0">
                    <div className="text-[10px] font-mono text-neutral-400">{e.event_number}</div>
                    <div className="text-sm font-bold text-neutral-900 truncate">{e.customer_name}</div>
                  </div>
                  <div className="shrink-0 text-right">
                    <div className="text-sm font-bold text-neutral-900">{won(e.total_amount)}</div>
                    <div className="text-[11px] text-neutral-400 flex items-center gap-0.5 justify-end">
                      {e.receive_method === 'visit' ? <><Store size={11} />매장</> : <><Truck size={11} />택배</>}
                    </div>
                  </div>
                </div>
                <div className="mt-1.5 text-xs text-neutral-500 truncate">
                  {(e.items || []).map((it) => `${it.product_name}${it.slicing ? '(슬라이싱)' : ''}×${it.qty}`).join(', ')}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>

      <SlidePanel open={!!sel} onClose={() => setSelId(null)} title="EVENT 접수 상세" className="sm:w-[440px]">
        {sel && <EventDetail ev={sel} patch={patch} del={del} onDone={() => setSelId(null)} goSales={() => router.push('/sales')} />}
      </SlidePanel>

      {editCampaign && <CampaignFormModal campaign={editCampaign} onClose={() => setEditCampaign(null)} create={createCampaign} update={updateCampaign} />}
    </>
  );
}

function CampaignFormModal({ campaign, onClose, create, update }: {
  campaign?: EventCampaign;
  onClose: () => void;
  create: ReturnType<typeof useCreateCampaign>;
  update: ReturnType<typeof useUpdateCampaign>;
}) {
  const isEdit = !!campaign;
  const [name, setName] = useState(campaign?.name || '');
  const [type, setType] = useState(campaign?.type || 'stock_clearance');
  const [status, setStatus] = useState(campaign?.status || 'active');
  const [rules, setRules] = useState<DiscountRule[]>(campaign?.discount_rules || []);
  const pending = create.isPending || update.isPending;

  const setRule = (i: number, k: keyof DiscountRule, v: number) =>
    setRules(rules.map((r, j) => (j === i ? { ...r, [k]: v } : r)));

  const save = () => {
    const cleanRules = rules.filter((r) => r.unit_price > 0 && r.min_qty > 0 && r.bundle_price > 0);
    if (isEdit) update.mutate({ id: campaign!.id, name: name.trim(), type, status, discount_rules: cleanRules }, { onSuccess: onClose });
    else create.mutate({ name: name.trim(), type, discount_rules: cleanRules }, { onSuccess: onClose });
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl p-5 w-full max-w-md max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <h3 className="text-base font-bold text-neutral-900 mb-4">{isEdit ? '캠페인 설정' : '새 캠페인'}</h3>

        <label className="text-xs text-neutral-500">캠페인명</label>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="예: 여름 한정 판매"
          className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm mb-3 mt-1" autoFocus />

        <div className="grid grid-cols-2 gap-2 mb-4">
          <div>
            <label className="text-xs text-neutral-500">유형</label>
            <select value={type} onChange={(e) => setType(e.target.value as EventCampaign['type'])} className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm mt-1">
              {Object.entries(CAMPAIGN_TYPE_LABEL).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
            </select>
          </div>
          {isEdit && (
            <div>
              <label className="text-xs text-neutral-500">상태</label>
              <select value={status} onChange={(e) => setStatus(e.target.value as EventCampaign['status'])} className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm mt-1">
                <option value="active">진행중</option>
                <option value="ended">종료</option>
              </select>
            </div>
          )}
        </div>

        {/* 묶음 할인 규칙 */}
        <div className="rounded-xl border border-neutral-200 p-3 mb-4">
          <div className="text-xs font-bold text-neutral-700 mb-1">묶음 할인 (같은 단가끼리)</div>
          <p className="text-[11px] text-neutral-400 mb-2">예: 단가 50000 · 3자루 · 묶음가 130000 → 3자루=13만, 4자루=18만, 6자루=26만</p>
          <div className="space-y-2">
            {rules.map((r, i) => (
              <div key={i} className="flex items-center gap-1.5">
                <input type="number" value={r.unit_price || ''} onChange={(e) => setRule(i, 'unit_price', parseInt(e.target.value) || 0)}
                  placeholder="단가" className="w-24 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-right" />
                <span className="text-[11px] text-neutral-400">원</span>
                <input type="number" value={r.min_qty || ''} onChange={(e) => setRule(i, 'min_qty', parseInt(e.target.value) || 0)}
                  placeholder="수량" className="w-14 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-right" />
                <span className="text-[11px] text-neutral-400">자루↑</span>
                <input type="number" value={r.bundle_price || ''} onChange={(e) => setRule(i, 'bundle_price', parseInt(e.target.value) || 0)}
                  placeholder="묶음가" className="flex-1 min-w-0 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-right" />
                <button onClick={() => setRules(rules.filter((_, j) => j !== i))} className="text-neutral-400 hover:text-red-500 shrink-0"><X size={14} /></button>
              </div>
            ))}
          </div>
          <button onClick={() => setRules([...rules, { unit_price: 0, min_qty: 3, bundle_price: 0 }])}
            className="mt-2 flex items-center gap-1 text-xs font-semibold text-indigo-600"><Plus size={13} />할인 규칙 추가</button>
        </div>

        <div className="flex gap-2">
          <button onClick={onClose} className="flex-1 py-2.5 rounded-lg border border-neutral-200 text-sm">취소</button>
          <button disabled={!name.trim() || pending} onClick={save}
            className="flex-1 py-2.5 rounded-lg bg-neutral-900 text-white text-sm font-bold disabled:opacity-50">
            {pending ? '저장 중...' : isEdit ? '저장' : '생성'}
          </button>
        </div>
      </div>
    </div>
  );
}

function EventDetail({ ev, patch, del, onDone, goSales }: {
  ev: EventSubmission;
  patch: ReturnType<typeof useEventPatch>;
  del: ReturnType<typeof useEventDelete>;
  onDone: () => void;
  goSales: () => void;
}) {
  const Row = ({ label, value }: { label: string; value: React.ReactNode }) => (
    <div className="flex justify-between gap-2 py-1.5 border-b border-neutral-50 text-sm">
      <span className="text-neutral-400 shrink-0 whitespace-nowrap">{label}</span>
      <span className="text-right text-neutral-800 min-w-0">{value}</span>
    </div>
  );

  return (
    <div className="space-y-4">
      <div>
        <div className="text-[11px] font-mono text-neutral-400">{ev.event_number}</div>
        <div className="text-lg font-bold text-neutral-900">{ev.customer_name}</div>
        <div className="text-sm text-neutral-500">{fmtPhone(ev.customer_phone)}</div>
        <span className="inline-block mt-1 text-[11px] px-2 py-0.5 rounded-full bg-neutral-100 text-neutral-600">{EVENT_STATUS_LABEL[ev.status]}</span>
      </div>

      <div className="rounded-xl bg-neutral-50 p-3 space-y-1">
        {(ev.items || []).map((it, i) => (
          <div key={i} className="flex justify-between text-sm">
            <span className="text-neutral-700 min-w-0 truncate">{it.product_name}{it.slicing ? ' (슬라이싱+2만)' : ''} ×{it.qty}</span>
            <span className="shrink-0 text-neutral-900">{won(it.unit_price * it.qty)}</span>
          </div>
        ))}
        {ev.slicing_addon > 0 && (
          <div className="flex justify-between text-sm text-neutral-500">
            <span>슬라이싱 가공</span><span>+{won(ev.slicing_addon)}</span>
          </div>
        )}
        <div className="flex justify-between pt-1 mt-1 border-t border-neutral-200 text-sm font-bold">
          <span>합계</span><span>{won(ev.total_amount)}</span>
        </div>
      </div>

      <div>
        <Row label="수령방법" value={ev.receive_method === 'visit' ? '매장 방문' : '택배 발송'} />
        {ev.receive_method !== 'visit' && <Row label="배송지" value={[ev.address1, ev.address2].filter(Boolean).join(' ') || '-'} />}
        {ev.memo && <Row label="메모" value={ev.memo} />}
      </div>

      {/* 액션 */}
      <div className="space-y-2 pt-2">
        {ev.status === 'received' && (
          <>
            {/* 2메시지 흐름: 접수완료 알림톡에 계좌 포함 → 신규접수에서 바로 입금확인 */}
            <button
              disabled={patch.isPending}
              onClick={() => {
                if (!window.confirm(`${ev.customer_name}님 입금을 확인하고 판매로 전환합니다. (재고가 차감됩니다)`)) return;
                patch.mutate({ id: ev.id, action: 'confirm_payment' }, { onSuccess: onDone });
              }}
              className="w-full py-2.5 rounded-lg bg-emerald-600 text-white text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
            >
              {patch.isPending ? <Loader2 size={16} className="animate-spin" /> : '입금확인 → 판매 전환'}
            </button>
            {/* (선택) 별도 입금안내 알림톡이 필요한 경우 */}
            <button
              disabled={patch.isPending}
              onClick={() => patch.mutate({ id: ev.id, action: 'payment_notice' }, { onSuccess: onDone })}
              className="w-full py-2 rounded-lg border border-neutral-300 text-neutral-600 text-xs font-semibold disabled:opacity-50 flex items-center justify-center gap-2"
            >
              <Package size={14} /> 입금안내 별도 발송 (선택)
            </button>
          </>
        )}
        {ev.status === 'payment_noticed' && (
          <button
            disabled={patch.isPending}
            onClick={() => {
              if (!window.confirm(`${ev.customer_name}님 입금을 확인하고 판매로 전환합니다. (재고가 차감됩니다)`)) return;
              patch.mutate({ id: ev.id, action: 'confirm_payment' }, { onSuccess: onDone });
            }}
            className="w-full py-2.5 rounded-lg bg-emerald-600 text-white text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
          >
            {patch.isPending ? <Loader2 size={16} className="animate-spin" /> : '입금확인 → 판매 전환'}
          </button>
        )}
        {ev.status === 'converted' && (
          <button onClick={goSales} className="w-full py-2.5 rounded-lg border border-neutral-300 text-sm font-semibold text-neutral-700">
            판매로 전환됨 — 판매관리에서 발송/수령 처리 →
          </button>
        )}

        {/* 접수 취소 (soft) — 취소 상태가 아니면 노출. converted는 판매 별도 안내 */}
        {ev.status !== 'cancelled' && (
          <button
            disabled={patch.isPending}
            onClick={() => {
              const msg = ev.status === 'converted'
                ? '이 접수를 취소 처리합니다.\n※ 판매 건은 판매관리에서 별도로 취소하세요 — 여기선 접수 기록만 취소됩니다.'
                : '이 접수를 취소합니다.';
              if (!window.confirm(msg)) return;
              patch.mutate({ id: ev.id, action: 'cancel' }, { onSuccess: onDone });
            }}
            className="w-full py-2 rounded-lg text-xs text-neutral-400 hover:text-red-500"
          >
            접수 취소
          </button>
        )}

        {/* 완전 삭제 (hard) — 취소/전환 건 정리용(오등록·테스트). 판매는 판매관리에서 별도 */}
        {(ev.status === 'cancelled' || ev.status === 'converted') && (
          <button
            disabled={del.isPending}
            onClick={() => {
              if (!window.confirm('이 접수 기록을 완전히 삭제합니다. 되돌릴 수 없습니다.\n※ 연결된 판매 건은 판매관리에서 별도 관리됩니다.')) return;
              del.mutate({ id: ev.id }, { onSuccess: onDone });
            }}
            className="w-full py-2 rounded-lg text-xs font-semibold text-red-500 border border-red-200 hover:bg-red-50 disabled:opacity-50 flex items-center justify-center gap-1.5"
          >
            {del.isPending ? <Loader2 size={13} className="animate-spin" /> : <X size={13} />}
            접수 기록 삭제
          </button>
        )}
      </div>
    </div>
  );
}
