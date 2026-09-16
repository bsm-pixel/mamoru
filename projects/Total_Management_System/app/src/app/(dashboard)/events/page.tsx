'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useEvents, useEventPatch, useEventDelete, useCampaigns, useCreateCampaign, useUpdateCampaign } from '@/hooks/use-events';
import { renderEventReceivedPreview, PREVIEW_SAMPLE } from '@/lib/event/campaign-notify';
import { EVENT_STATUS_LABEL, CAMPAIGN_TYPE_LABEL, type EventSubmission, type EventStatus, type EventCampaign, type DiscountRule } from '@/lib/event/types';
import { Zap, Loader2, Package, Truck, Store, ArrowLeft, Plus, ExternalLink, Settings, X } from 'lucide-react';
import { useIsLg } from '@/hooks/use-grid-mode';
import { useActivityTypes } from '@/hooks/use-activity-types';
import { ActivityChips } from '@/components/shared/activity-chips';
import { backdropClose } from '@/lib/ui/backdrop';
import { EscClose } from '@/components/ui/esc-close';

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
  const eventActTypes = useActivityTypes(list.map((e) => e.customer_phone));
  const tabCounts = countsByCampaign[campaignId || '_none'] || {};
  const sel = useMemo(() => (all || []).find((e) => e.id === selId) || null, [all, selId]);
  // 152: 선택한 접수의 캠페인이 무료(증정·체험단)인가 — 버튼 문구·확인 문구가 달라진다
  const selIsFree = useMemo(() => {
    const cid = (sel as { campaign_id?: string } | null | undefined)?.campaign_id;
    const c = (campaigns || []).find((x) => x.id === cid) as { payment_type?: string } | undefined;
    return c?.payment_type === 'free';
  }, [campaigns, sel]);
  const isLg = useIsLg();

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

        {(() => {
          const listContent = isLoading ? (
            <div className="py-16 text-center text-sm text-neutral-400"><Loader2 size={20} className="animate-spin inline" /></div>
          ) : list.length === 0 ? (
            <div className="py-16 text-center text-sm text-neutral-400">{EVENT_STATUS_LABEL[tab]} 건이 없습니다</div>
          ) : (
            <div className="space-y-2">
              {list.map((e) => (
                <button key={e.id} onClick={() => setSelId(e.id)}
                  className={`w-full text-left bg-white rounded-xl border px-4 py-3 transition ${selId === e.id ? 'border-neutral-900 ring-1 ring-neutral-900' : 'border-neutral-200 hover:border-neutral-400'}`}>
                  <div className="flex items-center justify-between gap-2">
                    <div className="min-w-0">
                      <div className="text-[10px] font-mono text-neutral-400">{e.event_number}</div>
                      <div className="flex items-center gap-1 min-w-0">
                        <span className="text-sm font-bold text-neutral-900 truncate">{e.customer_name}</span>
                        <ActivityChips types={eventActTypes(e.customer_phone)} className="shrink-0" />
                      </div>
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
          );
          // PC: 좌목록 + 우 상세(마스터-디테일) / 모바일: 목록만(상세는 SlidePanel)
          return isLg ? (
            <div className="flex gap-4 h-[calc(100vh-260px)]">
              <div className="flex-1 min-w-0 overflow-y-auto">{listContent}</div>
              <div className="w-[400px] shrink-0 overflow-y-auto">
                {sel ? (
                  <EventDetail ev={sel} patch={patch} del={del} onDone={() => setSelId(null)} goSales={(saleId?: string) => router.push(saleId ? `/sales/${saleId}` : '/sales')} isFree={selIsFree} />
                ) : (
                  <div className="flex flex-col items-center justify-center h-60 text-stone-400">
                    <Zap size={28} className="mb-2 opacity-40" />
                    <p className="text-xs text-center">접수를 선택하면<br />상세가 표시됩니다</p>
                  </div>
                )}
              </div>
            </div>
          ) : listContent;
        })()}
      </div>

      {!isLg && (
        <SlidePanel open={!!sel} onClose={() => setSelId(null)} title="EVENT 접수 상세" className="sm:w-[440px]">
          {sel && <EventDetail ev={sel} patch={patch} del={del} onDone={() => setSelId(null)} goSales={(saleId?: string) => router.push(saleId ? `/sales/${saleId}` : '/sales')} isFree={selIsFree} />}
        </SlidePanel>
      )}

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
  // 145: 알림톡 EVENT_신청완료의 안내 문구 — 비우면 기본 문구 발송
  const [notice, setNotice] = useState(campaign?.customer_notice || '');
  // 152: 무료(증정·체험단) 이벤트 + 신청 항목 표기
  const [paymentType, setPaymentType] = useState<'paid' | 'free'>(
    (campaign as { payment_type?: string } | undefined)?.payment_type === 'free' ? 'free' : 'paid',
  );
  const [itemsLabel, setItemsLabel] = useState((campaign as { items_label?: string } | undefined)?.items_label || '');
  const pending = create.isPending || update.isPending;
  const saveError = (update.error || create.error) as Error | null;

  const setRule = (i: number, k: keyof DiscountRule, v: number) =>
    setRules(rules.map((r, j) => (j === i ? { ...r, [k]: v } : r)));

  // 실시간 미리보기 — 발송 경로와 **같은 조립 함수**를 써서 화면과 실제 발송이 갈라지지 않게 한다
  const preview = renderEventReceivedPreview({
    event_name: name.trim() || '(캠페인명)',
    payment_type: paymentType,
    items_label: itemsLabel.trim(),
    notice_raw: notice.trim(),
  });

  const save = () => {
    const cleanRules = rules.filter((r) => r.unit_price > 0 && r.min_qty > 0 && r.bundle_price > 0);
    // 안내 문구는 바뀌었을 때만 전송 (마이그 145 실행 전에도 다른 설정 저장이 막히지 않게)
    const noticeChanged = notice.trim() !== (campaign?.customer_notice || '').trim();
    // 152 도 같은 규칙 — 바뀐 것만 전송해서 마이그 미실행 시에도 다른 설정 저장이 막히지 않게
    const prevPay = (campaign as { payment_type?: string } | undefined)?.payment_type === 'free' ? 'free' : 'paid';
    const prevLabel = (campaign as { items_label?: string } | undefined)?.items_label || '';
    const payChanged = paymentType !== prevPay;
    const labelChanged = itemsLabel.trim() !== prevLabel.trim();
    if (isEdit) update.mutate({
      id: campaign!.id, name: name.trim(), type, status, discount_rules: cleanRules,
      ...(noticeChanged ? { customer_notice: notice.trim() } : {}),
      ...(payChanged ? { payment_type: paymentType } : {}),
      ...(labelChanged ? { items_label: itemsLabel.trim() } : {}),
    }, { onSuccess: onClose });
    else create.mutate({
      name: name.trim(), type, discount_rules: cleanRules,
      ...(notice.trim() ? { customer_notice: notice.trim() } : {}),
      ...(paymentType === 'free' ? { payment_type: paymentType } : {}),
      ...(itemsLabel.trim() ? { items_label: itemsLabel.trim() } : {}),
    }, { onSuccess: onClose });
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4" {...backdropClose(onClose)}>
      <EscClose onClose={onClose} />
      <div className="bg-white rounded-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden" onClick={(e) => e.stopPropagation()}>
        <h3 className="text-base font-bold text-neutral-900 px-5 pt-5 pb-3 shrink-0">{isEdit ? '캠페인 설정' : '새 캠페인'}</h3>
        <div className="grid lg:grid-cols-[1fr_330px] flex-1 min-h-0 overflow-hidden">
        <div className="px-5 pb-5 overflow-y-auto">

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

        {/* 152: 결제 방식 — 무료면 알림톡에 금액·계좌 줄이 아예 생기지 않는다 */}
        <div className="rounded-xl border border-neutral-200 p-3 mb-4">
          <div className="text-xs font-bold text-neutral-700 mb-2">결제 방식</div>
          <div className="flex gap-2">
            {([['paid', '유료 (입금 받음)'], ['free', '무료 (증정·체험단)']] as const).map(([v, label]) => (
              <button key={v} type="button" onClick={() => setPaymentType(v)}
                className={`flex-1 py-2 rounded-lg text-xs font-semibold border transition ${paymentType === v ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-600 border-neutral-200 hover:bg-neutral-50'}`}>
                {label}
              </button>
            ))}
          </div>
          <p className="text-[11px] text-neutral-400 mt-2 leading-relaxed">
            {paymentType === 'free'
              ? '신청완료 알림톡에서 결제 금액·입금 계좌가 빠집니다. 진행도 [입금확인] 대신 [신청 확정]으로 바뀌고, 입금확인 알림톡은 보내지 않습니다.'
              : '신청완료 알림톡에 결제 금액과 입금 계좌가 안내됩니다.'}
          </p>
        </div>

        {/* 152: 신청 항목 표기 — 고객에게 보이는 줄만 바뀌고 내부 품목·재고는 그대로 */}
        <div className="rounded-xl border border-neutral-200 p-3 mb-4">
          <div className="text-xs font-bold text-neutral-700 mb-1">신청 항목 표기 <span className="font-normal text-neutral-400">(선택)</span></div>
          <p className="text-[11px] text-neutral-400 mb-2">알림톡 「신청 내역」에 제품명 대신 보여줄 문구. 비우면 <b>제품명 N개</b>가 그대로 나갑니다. 재고·판매 전환에는 영향 없습니다</p>
          <input value={itemsLabel} onChange={(e) => setItemsLabel(e.target.value)} maxLength={40}
            placeholder="예: 체험단 신청 / 증정 이벤트 응모"
            className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm" />
          <div className="text-right text-[10px] text-neutral-400 mt-1">{itemsLabel.length}/40</div>
        </div>

        {/* 145: 알림톡 고객 안내 문구 — EVENT_신청완료 🔔 진행 안내 첫 줄 */}
        <div className="rounded-xl border border-neutral-200 p-3 mb-4">
          <div className="text-xs font-bold text-neutral-700 mb-1">신청완료 알림톡 안내 문구</div>
          <p className="text-[11px] text-neutral-400 mb-2">이벤트별로 고객에게 따로 알릴 내용. <b>여러 줄</b>로 써도 됩니다. <b>할인·홍보 문구는 넣지 마세요</b> — 검수 반려 사유입니다</p>
          <textarea value={notice} onChange={(ev) => setNotice(ev.target.value)} maxLength={300} rows={4}
            placeholder={'예)\n보내실 가위는 신청 후 3일 안에 발송해 주세요\n받으실 주소가 바뀌면 이 채팅으로 회신해 주세요'}
            className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm leading-relaxed resize-y" />
          <div className="flex items-start justify-between gap-2 mt-1">
            <span className="text-[10px] text-neutral-400 leading-relaxed">
              {notice.trim()
                ? '→ 오른쪽 미리보기에 바로 반영됩니다'
                : paymentType === 'free'
                  ? '비워두면 → "바로 준비를 시작합니다" 한 줄만 나갑니다'
                  : '비워두면 → "입금이 확인되면 바로 준비를 시작합니다" 만 나갑니다'}
            </span>
            <span className="text-[10px] text-neutral-400 shrink-0">{notice.length}/300</span>
          </div>
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

        </div>

        {/* 실시간 미리보기 (2026-09-16)
            전엔 "비우면 기본 문구가 나갑니다"라고만 적혀 있고 그 문구가 화면에 없어서,
            사장님이 뭘 저장하는지 모르는 채 저장하게 됐다. 이제 눈으로 보고 정한다. */}
        <div className="border-t lg:border-t-0 lg:border-l border-neutral-200 bg-stone-50 px-4 py-4 overflow-y-auto">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold text-neutral-700">고객이 받을 알림톡</span>
            <span className="text-[10px] font-semibold text-neutral-400">신청완료</span>
          </div>
          <div className="rounded-2xl bg-[#B2C7DA] p-2.5">
            <div className="rounded-xl bg-white p-3 shadow-sm">
              <div className="text-[10px] font-bold text-neutral-400 mb-2 pb-2 border-b border-neutral-100">알림톡 도착</div>
              <pre className="whitespace-pre-wrap break-words font-sans text-[12px] leading-relaxed text-neutral-800 m-0">{preview}</pre>
              <div className="mt-2.5 pt-2 border-t border-neutral-100 text-center text-[11px] font-semibold text-neutral-500">1:1 문의</div>
            </div>
          </div>
          <p className="text-[10px] text-neutral-400 mt-2 leading-relaxed">
            예시 값 — {PREVIEW_SAMPLE.name}님 · {PREVIEW_SAMPLE.totalAmount.toLocaleString('ko-KR')}원 · 택배 발송.
            실제로는 접수한 고객의 이름·품목·금액·주소가 들어갑니다.
          </p>
          <p className="text-[10px] text-amber-600 mt-1.5 leading-relaxed">
            ⚠️ 솔라피에 <b>EVENT_신청완료</b> 새 본문이 등록·승인된 뒤부터 이 모양으로 나갑니다. 그 전에는 옛 템플릿이 발송됩니다.
          </p>
        </div>
        </div>

        {/* 푸터 — 모달 하단 고정. 좌측 컬럼 안에 두면 화면이 짧을 때 저장까지 스크롤해야 한다 */}
        <div className="shrink-0 border-t border-neutral-200 px-5 py-3 bg-white">
          {saveError && (
            <p className="text-xs text-red-500 mb-2">{(() => { try { return JSON.parse(saveError.message).error || saveError.message; } catch { return saveError.message; } })()}</p>
          )}
          <div className="flex gap-2">
            <button onClick={onClose} className="flex-1 py-2.5 rounded-lg border border-neutral-200 text-sm">취소</button>
            <button disabled={!name.trim() || pending} onClick={save}
              className="flex-1 py-2.5 rounded-lg bg-neutral-900 text-white text-sm font-bold disabled:opacity-50">
              {pending ? '저장 중...' : isEdit ? '저장' : '생성'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function EventDetail({ ev, patch, del, onDone, goSales, isFree }: {
  ev: EventSubmission;
  patch: ReturnType<typeof useEventPatch>;
  del: ReturnType<typeof useEventDelete>;
  onDone: () => void;
  goSales: (saleId?: string) => void;
  /** 152: 무료(증정·체험단) 캠페인 — "입금"이라는 말이 성립하지 않는다 */
  isFree: boolean;
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
            {/* 2메시지 흐름: 접수완료 알림톡에 계좌 포함 → 신규접수에서 바로 입금확인
                152: 무료 캠페인이면 "입금" 대신 "신청 확정" — 알림톡도 안 나간다 */}
            <button
              disabled={patch.isPending}
              onClick={() => {
                const msg = isFree
                  ? `${ev.customer_name}님 신청을 확정하고 발송 준비 단계로 넘깁니다. (재고가 차감됩니다)\n무료 이벤트라 입금확인 알림톡은 나가지 않습니다.`
                  : `${ev.customer_name}님 입금을 확인하고 판매로 전환합니다. (재고가 차감됩니다)`;
                if (!window.confirm(msg)) return;
                patch.mutate({ id: ev.id, action: 'confirm_payment' }, { onSuccess: onDone });
              }}
              className="w-full py-2.5 rounded-lg bg-emerald-600 text-white text-sm font-bold disabled:opacity-50 flex items-center justify-center gap-2"
            >
              {patch.isPending ? <Loader2 size={16} className="animate-spin" /> : (isFree ? '신청 확정 → 발송 준비' : '입금확인 → 판매 전환')}
            </button>
            {/* (선택) 별도 입금안내 알림톡 — 무료 이벤트엔 보낼 이유가 없어 숨긴다 */}
            {!isFree && (
              <button
                disabled={patch.isPending}
                onClick={() => patch.mutate({ id: ev.id, action: 'payment_notice' }, { onSuccess: onDone })}
                className="w-full py-2 rounded-lg border border-neutral-300 text-neutral-600 text-xs font-semibold disabled:opacity-50 flex items-center justify-center gap-2"
              >
                <Package size={14} /> 입금안내 별도 발송 (선택)
              </button>
            )}
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
          <button onClick={() => goSales(ev.sale_id || undefined)} className="w-full py-2.5 rounded-lg border border-neutral-300 text-sm font-semibold text-neutral-700">
            판매로 전환됨 — {ev.sale_id ? '전환된 판매 상세로 →' : '판매관리에서 발송/수령 처리 →'}
          </button>
        )}

        {/* 접수 취소 (soft) — 취소 상태가 아니면 노출. converted는 판매 별도 안내 */}
        {ev.status !== 'cancelled' && (
          <button
            disabled={patch.isPending}
            onClick={() => {
              const msg = ev.status === 'converted'
                ? '이 접수는 판매로 전환되었습니다.\n판매가 아직 살아있으면 취소가 막힙니다 — 판매관리에서 판매를 먼저 취소/반품하세요.\n계속하시겠습니까?'
                : '이 접수를 취소합니다.';
              if (!window.confirm(msg)) return;
              patch.mutate({ id: ev.id, action: 'cancel' }, {
                onSuccess: onDone,
                onError: (e) => window.alert(e instanceof Error ? e.message : '취소 실패'),
              });
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
