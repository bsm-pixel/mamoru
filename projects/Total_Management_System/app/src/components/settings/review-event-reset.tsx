'use client';

/**
 * 설정 > 시스템 「리뷰 추첨 초기화」 (2026-10-02)
 *
 * 추첨 화면(녹화·화면공유 대상)에는 "다시 뽑기"류 버튼을 두지 않는다 — 조작으로 오해받는다.
 * 연습 추첨을 지우는 일은 이 설정 화면에서만 한다. 판정은 전부 서버(api/reviews/event/reset)가 한다.
 */

import { useCallback, useEffect, useState } from 'react';
import toast from 'react-hot-toast';

interface Preview {
  month: string;
  status: string;
  winners: number;
  byRank: Record<string, number>;
  shipments: number;
  notified: number;
  addressed: number;
  invoiced: number;
  blocked: string | null;
}

const STATUS_LABEL: Record<string, string> = { draft: '비공개', live: '진행중', announced: '발표됨', none: '설정 없음' };

/** 지난달 (추첨은 보통 마감된 지난달 것) — KST 기준 */
function defaultMonth(): { y: number; m: number } {
  const now = new Date(Date.now() + 9 * 60 * 60 * 1000);
  const d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1));
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1 };
}

export default function ReviewEventReset() {
  const init = defaultMonth();
  const [year, setYear] = useState(init.y);
  const [mon, setMon] = useState(init.m);
  const [pv, setPv] = useState<Preview | null>(null);
  const [loading, setLoading] = useState(false);
  const [busy, setBusy] = useState(false);

  const yymm = `${String(year).slice(2)}${String(mon).padStart(2, '0')}`;
  const title = `${year}년 ${mon}월`;

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch(`/api/reviews/event/reset?month=${yymm}`);
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || '조회 실패');
      setPv(json);
    } catch (e) {
      setPv(null);
      toast.error(e instanceof Error ? e.message : '조회 실패');
    } finally {
      setLoading(false);
    }
  }, [yymm]);

  useEffect(() => { load(); }, [load]);

  async function reset() {
    if (!pv || pv.blocked || pv.winners === 0) return;
    const ranks = Object.keys(pv.byRank).sort().map((r) => `${r}등 ${pv.byRank[r]}명`).join(' · ');
    if (!window.confirm(`${title} 리뷰 이벤트 추첨을 초기화합니다.\n\n당첨 표시 ${pv.winners}명 (${ranks}) 이 모두 해제됩니다.\n상품·마감일 설정은 그대로 둡니다.\n\n계속할까요?`)) return;
    setBusy(true);
    try {
      const res = await fetch('/api/reviews/event/reset', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ month: yymm, confirm: yymm }),
      });
      const json = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(json.error || '초기화 실패');
      toast.success(`${title} 추첨을 초기화했습니다 (${json.cleared}명 해제)`);
      await load();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : '초기화 실패');
    } finally {
      setBusy(false);
    }
  }

  const thisYear = new Date().getFullYear();
  const years = [thisYear - 1, thisYear, thisYear + 1];
  const canReset = !!pv && !pv.blocked && pv.winners > 0;

  return (
    <div className="pt-6 border-t border-neutral-200">
      <h3 className="text-base font-bold mb-1">리뷰 추첨 초기화</h3>
      <p className="text-xs text-neutral-400 mb-4 leading-relaxed">
        연습으로 돌린 리뷰 이벤트 추첨을 지우고 처음 상태로 되돌립니다. 상품·마감일 설정은 그대로 둡니다.
        <br />이미 발표했거나 당첨자에게 안내·배송이 시작된 달은 초기화되지 않습니다.
      </p>

      <div className="flex flex-wrap items-center gap-2 mb-3">
        <select value={year} onChange={(e) => setYear(Number(e.target.value))}
          className="h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm">
          {years.map((y) => <option key={y} value={y}>{y}년</option>)}
        </select>
        <select value={mon} onChange={(e) => setMon(Number(e.target.value))}
          className="h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm">
          {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => <option key={m} value={m}>{m}월</option>)}
        </select>
        <span className="text-xs text-neutral-400">이벤트 달 (후기 마감 기준)</span>
      </div>

      <div className="rounded-lg border border-neutral-200 bg-neutral-50 px-4 py-3 text-sm">
        {loading ? (
          <p className="text-neutral-400">불러오는 중…</p>
        ) : !pv ? (
          <p className="text-neutral-400">불러오지 못했습니다</p>
        ) : (
          <>
            <p className="text-neutral-800">
              <b>{title}</b> · 상태 <b>{STATUS_LABEL[pv.status] || pv.status}</b> · 당첨 표시 <b>{pv.winners}명</b>
              {pv.winners > 0 && (
                <span className="text-neutral-500"> ({Object.keys(pv.byRank).sort().map((r) => `${r}등 ${pv.byRank[r]}`).join(' · ')})</span>
              )}
            </p>
            {pv.shipments > 0 && (
              <p className="text-xs text-neutral-500 mt-1">
                당첨자 배송: 안내 발송 {pv.notified} · 주소 입력 {pv.addressed} · 송장 {pv.invoiced}
              </p>
            )}
            {pv.blocked && <p className="text-xs text-amber-700 mt-2 leading-relaxed">초기화할 수 없습니다 — {pv.blocked}</p>}
            {!pv.blocked && pv.winners === 0 && <p className="text-xs text-neutral-400 mt-2">초기화할 당첨자가 없습니다.</p>}
          </>
        )}
      </div>

      {/* 파괴적 액션 — 평소엔 조용한 회색, hover 에서만 빨강 (액션 영역 표준) */}
      <button type="button" onClick={reset} disabled={!canReset || busy}
        className="mt-3 text-xs text-neutral-400 underline underline-offset-2 hover:text-red-600 disabled:opacity-40 disabled:no-underline disabled:hover:text-neutral-400">
        {busy ? '초기화 중…' : `${title} 추첨 초기화`}
      </button>
    </div>
  );
}
