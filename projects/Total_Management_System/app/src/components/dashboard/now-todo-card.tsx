'use client';

/**
 * 「지금 처리할 것」 (2026-11-15) — 대시보드 맨 위
 *
 * 왜: 숫자 카드만으로는 "그래서 지금 무엇부터 하면 되는지"를 알 수 없어 메뉴를 하나씩 열어 봐야 했다.
 *   신규 상담 · 처리 대기 판매 · 미수금 · 매입 대기 · 확인 대기 경비를 한곳에 줄로 보여 주고,
 *   줄을 누르면 그 건이 바로 열린다(종류마다 3건까지, 나머지는 "외 N건"으로 목록 이동).
 * 처리할 것이 하나도 없으면 한 줄만 남는다.
 */

import Link from 'next/link';
import { useQuery } from '@tanstack/react-query';
import { ChevronRight, CircleCheck } from 'lucide-react';
import { useConsultations } from '@/hooks/use-consultations';
import { useSales } from '@/hooks/use-sales';
import { useOutstandingAlert, usePurchasingAlert } from '@/hooks/use-dashboard-stats';
import { formatDate, formatKRW, CONSULTATION_TYPE_LABEL } from '@/lib/utils/format';
import { TONE, type StatusTone } from '@/lib/ui/status-tone';

const MAX_ROWS = 3;

interface Row { key: string; href: string; title: string; sub?: string; right?: string }
interface Group { key: string; label: string; tone: StatusTone; total: number; listHref: string; rows: Row[] }

export function NowTodoCard() {
  const { data: consult } = useConsultations({ statuses: ['pending_admin', 'reschedule_requested', 'change_requested'], limit: MAX_ROWS });
  const { data: sales } = useSales({ tab: 'processing', dateRange: 'all', limit: MAX_ROWS });
  const { data: outstanding } = useOutstandingAlert();
  const { data: purchasing } = usePurchasingAlert();
  // 경비 화면의 「확인 대기」와 같은 자료(같은 queryKey → 한 번만 받는다)
  const { data: inbox } = useQuery({
    queryKey: ['expense-inbox'],
    queryFn: async () => {
      const res = await fetch('/api/expenses/inbox');
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ items: Array<{ id: string }> }>;
    },
  });

  const groups: Group[] = [
    {
      key: 'consult', label: '신규 상담', tone: 'wait', total: consult?.total || 0, listHref: '/consultations',
      rows: (consult?.consultations || []).slice(0, MAX_ROWS).map((c) => ({
        key: c.id, href: `/consultations/${c.id}`, title: c.name || '이름 없음',
        sub: `${CONSULTATION_TYPE_LABEL[c.consultation_type] || '상담'}${c.status !== 'pending_admin' ? ' · 일정 변경 요청' : ''}`,
        right: c.received_at ? formatDate(c.received_at) : undefined,
      })),
    },
    {
      key: 'sales', label: '처리 대기 판매', tone: 'wait', total: sales?.total || 0, listHref: '/sales',
      rows: (sales?.sales || []).slice(0, MAX_ROWS).map((s) => ({
        key: s.id, href: `/sales/${s.id}`, title: s.customer_name || '고객 미지정',
        sub: `${s.sale_number || ''} · 결제완료 · 출고 전`, right: formatDate(s.sale_date),
      })),
    },
    {
      key: 'outstanding', label: '미수금', tone: 'problem', total: outstanding?.length || 0, listHref: '/customers',
      rows: (outstanding || []).slice(0, MAX_ROWS).map((c) => ({
        key: c.id, href: `/customers/${c.id}`, title: c.name, right: formatKRW(c.outstanding_balance),
      })),
    },
    {
      key: 'purchasing', label: '매입 대기', tone: 'progress', total: purchasing?.length || 0, listHref: '/purchasing',
      rows: (purchasing || []).slice(0, MAX_ROWS).map((p) => ({
        key: p.id, href: `/purchasing/${p.id}`, title: p.supplier_name || p.po_number,
        sub: p.po_number, right: p.expected_date ? `입고 ${formatDate(p.expected_date)}` : '입고일 미정',
      })),
    },
    {
      key: 'expense', label: '확인 대기 경비', tone: 'wait', total: inbox?.items?.length || 0, listHref: '/expenses',
      rows: (inbox?.items?.length || 0) > 0
        ? [{ key: 'inbox', href: '/expenses', title: `메일에서 찾은 영수증 ${inbox!.items.length}건`, sub: '경비로 넣을지 확인' }]
        : [],
    },
  ].filter((g) => g.total > 0) as Group[];

  const totalCount = groups.reduce((n, g) => n + g.total, 0);

  if (groups.length === 0) {
    return (
      <div className="bg-white rounded-2xl border border-stone-200 px-4 py-3 flex items-center gap-2 text-sm text-stone-500">
        <CircleCheck size={16} className="text-stone-400" />
        지금 처리할 것이 없습니다
      </div>
    );
  }

  return (
    <section aria-label="지금 처리할 것" className="bg-white rounded-2xl border border-stone-200 p-4">
      <div className="flex items-baseline gap-2 mb-3">
        <h3 className="text-sm font-bold text-stone-900">지금 처리할 것</h3>
        <span className="text-xs font-semibold text-stone-500 tabular-nums">{totalCount}건</span>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-x-6 gap-y-4">
        {groups.map((g) => (
          <div key={g.key} className="min-w-0">
            <Link href={g.listHref} className="flex items-center gap-1.5 mb-1.5 group">
              <span className={`px-1.5 py-0.5 rounded text-[10.5px] font-bold ${TONE[g.tone]}`}>{g.label}</span>
              <span className="text-xs font-bold text-stone-800 tabular-nums">{g.total}</span>
              <ChevronRight size={12} className="text-stone-300 group-hover:text-stone-600 transition" />
            </Link>
            <ul className="divide-y divide-stone-100 border-t border-stone-100">
              {g.rows.map((r) => (
                <li key={r.key}>
                  <Link href={r.href} className="flex items-center justify-between gap-3 py-2 px-1 -mx-1 rounded hover:bg-stone-50 transition">
                    <span className="min-w-0">
                      <span className="block text-[13px] font-semibold text-stone-800 truncate">{r.title}</span>
                      {r.sub && <span className="block text-[11px] text-stone-400 truncate">{r.sub}</span>}
                    </span>
                    {r.right && <span className="text-xs text-stone-500 tabular-nums shrink-0">{r.right}</span>}
                  </Link>
                </li>
              ))}
              {g.total > g.rows.length && g.key !== 'expense' && (
                <li>
                  <Link href={g.listHref} className="block py-1.5 px-1 text-[11px] font-semibold text-stone-500 hover:text-stone-900">
                    외 {g.total - g.rows.length}건 모두 보기
                  </Link>
                </li>
              )}
            </ul>
          </div>
        ))}
      </div>
    </section>
  );
}
