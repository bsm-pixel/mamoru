'use client';

/**
 * 리뷰 이벤트 「당첨자 배송」 (2026-10-02, 마이그 156)
 *
 * 당첨자 = 선정 화면에서 저장된 reviews.event_rank 그대로. 이 패널은 배송만 다룬다:
 *   [당첨 안내 보내기](배송지 입력 링크 알림톡) → 고객 주소 저장(또는 [직접 입력]) → [송장 생성] → 기사님 집하 → 출고 알림톡(자동)
 * 템플릿 승인 전에는 [링크 복사]로 카톡 채널 1:1 채팅에서 직접 보내면 같은 흐름으로 이어진다.
 */

import { useCallback, useEffect, useState } from 'react';
import { Truck, Send, Link2, Loader2, MapPin, RotateCw, X, PackageCheck } from 'lucide-react';
import toast from 'react-hot-toast';
import { DaumPostcodeButton } from '@/components/shared/daum-postcode-button';
import { EscClose } from '@/components/ui/esc-close';

interface Prefill { postcode: string; address_road: string; address_detail: string; source: string }
interface Item {
  id: string; rank: number | null; rank_label: string | null; prize: string | null;
  name: string; phone: string; address_link: string;
  postcode: string | null; address_road: string | null; address_detail: string | null; delivery_message: string | null;
  address_submitted_at: string | null; won_notified_at: string | null;
  invoice_number: string | null; invoice_created_at: string | null;
  shipped_at: string | null; shipped_notified_at: string | null; delivered_at: string | null;
  prefill: Prefill | null;
}

const SRC: Record<string, string> = { customer: '고객정보', repair: '복원수리 접수', order: '아임웹 주문' };

function fmt(iso: string | null): string {
  if (!iso) return '';
  const d = new Date(iso);
  return `${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}
function phoneFmt(p: string): string {
  return p.length === 11 ? `${p.slice(0, 3)}-${p.slice(3, 7)}-${p.slice(7)}` : p;
}

export default function EventShipmentsPanel({ month }: { month: string }) {
  const [items, setItems] = useState<Item[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState<string | null>(null);
  const [editing, setEditing] = useState<Item | null>(null);
  const [form, setForm] = useState({ postcode: '', address_road: '', address_detail: '', delivery_message: '' });

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch(`/api/reviews/event/shipments?month=${month}`);
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || '조회 실패');
      setItems(json.items || []);
    } catch (e) {
      toast.error(e instanceof Error ? e.message : '당첨자 배송 조회 실패');
    } finally {
      setLoading(false);
    }
  }, [month]);

  useEffect(() => { load(); }, [load]);

  async function post(body: Record<string, unknown>, key: string) {
    setBusy(key);
    try {
      const res = await fetch('/api/reviews/event/shipments', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
      });
      const json = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(json.error || '처리 실패');
      return json;
    } catch (e) {
      toast.error(e instanceof Error ? e.message : '처리 실패');
      return null;
    } finally {
      setBusy(null);
    }
  }

  async function notify(ids?: string[]) {
    const n = ids ? ids.length : items.filter((i) => !i.won_notified_at).length;
    if (!n) { toast('보낼 대상이 없습니다'); return; }
    if (!window.confirm(`당첨 안내 알림톡을 ${n}명에게 보냅니다.\n(솔라피 템플릿 승인 + Make 연결 전이면 고객에게 도착하지 않습니다)`)) return;
    const r = await post({ action: 'notify', month, ids }, ids ? `notify-${ids[0]}` : 'notify-all');
    if (r) {
      if (r.failed?.length) toast.error(`${r.sent}명 발송 · 실패 ${r.failed.length}명: ${r.failed.join(', ')}`);
      else toast.success(`${r.sent}명에게 당첨 안내를 보냈습니다`);
      load();
    }
  }

  async function copyLink(it: Item) {
    const text = `${it.name}님, MAMORU 리뷰 이벤트 ${it.rank_label || ''} 당첨을 축하드립니다\n당첨 상품: ${it.prize || ''}\n아래 링크에서 받으실 주소를 확인해 주세요\n${it.address_link}`;
    try {
      await navigator.clipboard.writeText(text);
      toast.success('안내 문구 + 링크를 복사했습니다 (카톡 채널 채팅에 붙여넣기)');
    } catch {
      window.prompt('복사해서 보내주세요', it.address_link);
    }
  }

  async function makeInvoice(it: Item) {
    if (!window.confirm(`${it.name}님 송장을 생성합니다.\n${[it.address_road, it.address_detail].filter(Boolean).join(' ')}`)) return;
    const r = await post({ action: 'invoice', id: it.id }, `inv-${it.id}`);
    if (r) { toast.success(`송장 ${r.invoiceNumber} 생성 — ALPS에서 출력하세요`); load(); }
  }

  async function cancelInvoice(it: Item) {
    if (!window.confirm(`송장 ${it.invoice_number} 을(를) 취소합니다.`)) return;
    const r = await post({ action: 'cancel_invoice', id: it.id }, `cinv-${it.id}`);
    if (r) { toast.success('송장을 취소했습니다'); load(); }
  }

  function openEdit(it: Item) {
    const src = it.address_road ? it : it.prefill;
    setForm({
      postcode: src?.postcode || '', address_road: src?.address_road || '',
      address_detail: src?.address_detail || '', delivery_message: it.delivery_message || '',
    });
    setEditing(it);
  }

  async function saveAddress() {
    if (!editing) return;
    const r = await post({ action: 'save_address', id: editing.id, ...form }, 'addr');
    if (r) { toast.success('배송지를 저장했습니다 (고객정보 주소도 갱신)'); setEditing(null); load(); }
  }

  const total = items.length;
  const addrDone = items.filter((i) => i.address_submitted_at).length;
  const invDone = items.filter((i) => i.invoice_number).length;
  const shipDone = items.filter((i) => i.shipped_at).length;
  const notNotified = items.filter((i) => !i.won_notified_at).length;

  return (
    <div className="mt-6 rounded-xl border border-stone-200 bg-white">
      <div className="flex flex-wrap items-center gap-2 px-4 py-3 border-b border-stone-100">
        <Truck size={16} className="text-stone-700" />
        <span className="text-sm font-bold text-stone-900">당첨자 배송</span>
        {total > 0 && (
          <span className="text-xs text-stone-500">
            배송지 {addrDone}/{total} · 송장 {invDone} · 출고 {shipDone}
          </span>
        )}
        <button type="button" onClick={load} className="ml-auto p-1.5 rounded-md text-stone-400 hover:text-stone-700 hover:bg-stone-50" title="새로고침">
          <RotateCw size={14} className={loading ? 'animate-spin' : ''} />
        </button>
        <button type="button" disabled={!!busy || notNotified === 0} onClick={() => notify()}
          className="px-3 py-1.5 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 disabled:opacity-40 flex items-center gap-1.5">
          {busy === 'notify-all' ? <Loader2 size={13} className="animate-spin" /> : <Send size={13} />}
          당첨 안내 보내기{notNotified > 0 ? ` (${notNotified}명)` : ''}
        </button>
      </div>

      {loading ? (
        <p className="px-4 py-6 text-sm text-stone-400">불러오는 중…</p>
      ) : total === 0 ? (
        <p className="px-4 py-6 text-sm text-stone-400">이 달에 저장된 당첨자가 없습니다. 위에서 선정 후 [저장]하면 여기에 나타납니다.</p>
      ) : (
        <ul className="divide-y divide-stone-100">
          {items.map((it) => {
            const addr = [it.address_road, it.address_detail].filter(Boolean).join(' ');
            return (
              <li key={it.id} className="px-4 py-3 flex flex-col lg:flex-row lg:items-center gap-2 lg:gap-4">
                {/* 1. 누구 · 무엇 */}
                <div className="lg:w-56 shrink-0 min-w-0">
                  <div className="flex items-center gap-1.5">
                    <span className="text-[11px] font-semibold px-1.5 py-0.5 rounded bg-stone-100 text-stone-600 shrink-0">{it.rank_label || `${it.rank}등`}</span>
                    <span className="text-sm font-semibold text-stone-900 truncate">{it.name}</span>
                  </div>
                  <div className="text-xs text-stone-500 mt-0.5 truncate">{phoneFmt(it.phone)} · {it.prize || '상품 미입력'}</div>
                </div>

                {/* 2. 배송지 */}
                <div className="flex-1 min-w-0 text-xs">
                  {it.address_submitted_at ? (
                    <div className="flex items-start gap-1.5">
                      <MapPin size={13} className="text-emerald-600 mt-0.5 shrink-0" />
                      <span className="text-stone-700 break-keep">{addr}{it.delivery_message ? ` · ${it.delivery_message}` : ''}</span>
                    </div>
                  ) : (
                    <div className="text-stone-400">
                      배송지 미입력
                      {it.prefill && <span className="block text-stone-400 truncate">예전 주소({SRC[it.prefill.source] || it.prefill.source}): {it.prefill.address_road}</span>}
                    </div>
                  )}
                </div>

                {/* 3. 알림 · 링크 */}
                <div className="flex items-center gap-1.5 shrink-0 text-xs">
                  {it.won_notified_at
                    ? <span className="text-stone-500">안내 {fmt(it.won_notified_at)}</span>
                    : <button type="button" disabled={!!busy} onClick={() => notify([it.id])}
                        className="px-2 py-1 rounded-md border border-stone-200 text-stone-700 hover:bg-stone-50 disabled:opacity-40 flex items-center gap-1">
                        {busy === `notify-${it.id}` ? <Loader2 size={12} className="animate-spin" /> : <Send size={12} />}안내
                      </button>}
                  <button type="button" onClick={() => copyLink(it)}
                    className="px-2 py-1 rounded-md border border-stone-200 text-stone-700 hover:bg-stone-50 flex items-center gap-1" title="안내 문구 + 배송지 입력 링크 복사">
                    <Link2 size={12} />링크
                  </button>
                  {!it.invoice_number && (
                    <button type="button" onClick={() => openEdit(it)}
                      className="px-2 py-1 rounded-md border border-stone-200 text-stone-700 hover:bg-stone-50">
                      {it.address_submitted_at ? '주소 수정' : '직접 입력'}
                    </button>
                  )}
                </div>

                {/* 4. 송장 · 출고 */}
                <div className="flex items-center gap-1.5 shrink-0 text-xs lg:w-52 lg:justify-end">
                  {it.delivered_at ? (
                    <span className="flex items-center gap-1 text-emerald-700 font-semibold"><PackageCheck size={13} />배달완료 · {it.invoice_number}</span>
                  ) : it.shipped_at ? (
                    <span className="text-blue-700 font-semibold">출고 {fmt(it.shipped_at)} · {it.invoice_number}{it.shipped_notified_at ? ' · 알림 ✓' : ''}</span>
                  ) : it.invoice_number ? (
                    <>
                      <span className="text-stone-700 font-mono">{it.invoice_number}</span>
                      <span className="text-stone-400">집하 대기</span>
                      <button type="button" disabled={!!busy} onClick={() => cancelInvoice(it)}
                        className="p-1 rounded text-stone-400 hover:text-red-600 disabled:opacity-40" title="송장 취소"><X size={13} /></button>
                    </>
                  ) : (
                    <button type="button" disabled={!!busy || !it.address_road} onClick={() => makeInvoice(it)}
                      className="px-2.5 py-1 rounded-md bg-stone-900 text-white font-semibold hover:bg-stone-800 disabled:opacity-30 flex items-center gap-1"
                      title={it.address_road ? '롯데 송장 생성' : '배송지가 있어야 송장을 만들 수 있습니다'}>
                      {busy === `inv-${it.id}` ? <Loader2 size={12} className="animate-spin" /> : <Truck size={12} />}송장 생성
                    </button>
                  )}
                </div>
              </li>
            );
          })}
        </ul>
      )}

      <p className="px-4 py-3 text-[11px] text-stone-400 leading-relaxed border-t border-stone-100">
        · <b>당첨 안내</b> = 배송지 입력 링크가 담긴 알림톡. 솔라피 템플릿 승인 전에는 <b>[링크]</b>로 복사해 카톡 채널 채팅에서 보내도 같은 흐름입니다.<br />
        · 고객이 주소를 저장하면 앱 알림이 오고, <b>고객정보 주소도 새 주소로 갱신</b>됩니다. 전화로 받은 주소는 <b>[직접 입력]</b>.<br />
        · <b>[송장 생성]</b> 후 ALPS에서 출력 → 기사님 집하 스캔 시 <b>출고 알림톡이 자동</b>으로 나갑니다(30분 주기 확인).
      </p>

      {/* 배송지 직접 입력/수정 */}
      {editing && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
          <EscClose onClose={() => setEditing(null)} />
          <div className="w-full max-w-md rounded-xl bg-white p-5 shadow-xl">
            <div className="flex items-center justify-between mb-3">
              <p className="text-sm font-bold text-stone-900">{editing.name}님 배송지</p>
              <button type="button" onClick={() => setEditing(null)} className="p-1 text-stone-400 hover:text-stone-700"><X size={16} /></button>
            </div>
            <div className="flex gap-2 mb-2">
              <input value={form.postcode} readOnly placeholder="우편번호"
                className="w-24 h-9 px-3 rounded-lg border border-stone-200 bg-stone-50 text-sm text-stone-600" />
              <DaumPostcodeButton onSelected={(d) => setForm((f) => ({ ...f, postcode: d.zonecode, address_road: d.roadAddress }))}>
                주소검색
              </DaumPostcodeButton>
            </div>
            <input value={form.address_road} readOnly placeholder="도로명 주소 (주소검색으로 입력)"
              className="w-full h-9 px-3 mb-2 rounded-lg border border-stone-200 bg-stone-50 text-sm text-stone-600" />
            <input value={form.address_detail} onChange={(e) => setForm((f) => ({ ...f, address_detail: e.target.value }))} placeholder="상세 주소"
              className="w-full h-9 px-3 mb-2 rounded-lg border border-stone-200 text-sm" />
            <input value={form.delivery_message} onChange={(e) => setForm((f) => ({ ...f, delivery_message: e.target.value }))} placeholder="배송 메모 (선택)"
              className="w-full h-9 px-3 mb-4 rounded-lg border border-stone-200 text-sm" />
            <button type="button" disabled={busy === 'addr' || !form.postcode || !form.address_road} onClick={saveAddress}
              className="w-full h-10 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 disabled:opacity-40">
              {busy === 'addr' ? '저장 중…' : '저장 (고객정보 주소도 갱신)'}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
