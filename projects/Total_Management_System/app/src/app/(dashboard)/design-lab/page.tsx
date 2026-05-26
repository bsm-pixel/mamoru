'use client';

/**
 * /design-lab — TMS 디자인 모니터 (사장님 + 클로드 협업 도구)
 *
 * 2026-05-26 Phase G-6: /sales/design-preview 에서 /design-lab 로 이동.
 *   - 향후 다른 페이지 디자인 검토 시 § 추가하여 한 화면에서 비교
 *   - 사이드바 메뉴 미등록 (URL 직접 접근 — 운영 메뉴 정돈 유지)
 *   - 운영 데이터 호출 X (모두 하드코딩 임시)
 *
 * 사용법:
 *   사장님: 새 디자인 검토 필요 시 "디자인 모니터에 § XXX 페이지 추가해줘" 한 마디
 *   클로드: 해당 § 섹션 추가 + 비교 옵션 3~4개 렌더 → 사장님 진입 → 결정 → 적용
 */

import { Topbar } from '@/components/layout/topbar';
import { formatKRW } from '@/lib/utils/format';

// ─── 임시 데이터 ───────────────────────────────────────────────
const STATS = {
  customer: {
    week: { amount: 1190000, count: 4 },
    month: { amount: 6567010, count: 27 },
    outstanding: 280000,
  },
  partner: {
    week: { amount: 2400000, count: 3 },
    month: { amount: 9800000, count: 15 },
    outstanding: 1500000,
  },
};

type RowState = 'paid_done' | 'paid_shipping' | 'paid_wait_ship' | 'unpaid' | 'partial' | 'shipped_b2b_unpaid' | 'cancelled';

type SampleRow = {
  id: string;
  sourceType: 'sale' | 'delivery';
  state: RowState;
  name: string;
  date: string;
  amount: number;
  paymentMethod: string;
  channelOrType: string;
  note?: string;
};

const SAMPLE_ROWS: SampleRow[] = [
  // 판매완료 (초록 줄)
  { id: '1', sourceType: 'sale', state: 'paid_done', name: '서윤 점장님 (박미애)', date: '5월 26일', amount: 100000, paymentMethod: '카드', channelOrType: '온라인상담', note: '배송완료 → 판매완료 자동 전환' },
  { id: '2', sourceType: 'sale', state: 'paid_done', name: '김매장 (매장수령)', date: '5월 25일', amount: 240000, paymentMethod: '카드', channelOrType: '오프라인', note: '매장 직접수령 (배송 X, 즉시 판매완료)' },
  { id: '3', sourceType: 'delivery', state: 'paid_done', name: '○○매장', date: '5월 24일', amount: 1240000, paymentMethod: '이체', channelOrType: '딜러', note: 'B2B 출고완료 + 결제완료 = 판매완료' },
  // 진행 중
  { id: '4', sourceType: 'sale', state: 'paid_shipping', name: '전아연', date: '5월 25일', amount: 780000, paymentMethod: '이체', channelOrType: '온라인상담', note: '결제완료 · 배송중' },
  { id: '5', sourceType: 'sale', state: 'paid_wait_ship', name: '곽경진', date: '5월 25일', amount: 60000, paymentMethod: '이체', channelOrType: '오프라인', note: '결제완료 · 송장발급 대기' },
  { id: '6', sourceType: 'delivery', state: 'shipped_b2b_unpaid', name: '△△아카데미', date: '5월 23일', amount: 580000, paymentMethod: '이체', channelOrType: '아카데미', note: 'B2B 출고완료 · 결제 대기' },
  // 미수금
  { id: '7', sourceType: 'sale', state: 'unpaid', name: '김승한', date: '5월 18일', amount: 280000, paymentMethod: '카드', channelOrType: '오프라인', note: '미결제 — 가장 시급' },
  { id: '8', sourceType: 'delivery', state: 'unpaid', name: '□□아카데미', date: '5월 20일', amount: 1500000, paymentMethod: '이체', channelOrType: '아카데미', note: '거래처 미결제 (큰 금액)' },
  // 부분결제
  { id: '9', sourceType: 'sale', state: 'partial', name: '박상민', date: '5월 22일', amount: 350000, paymentMethod: '카드', channelOrType: '온라인상담', note: '부분결제 (선납금만)' },
  // 취소
  { id: '10', sourceType: 'sale', state: 'cancelled', name: '취소건 예시', date: '5월 20일', amount: 150000, paymentMethod: '카드', channelOrType: '오프라인', note: '취소 — 흐림 표시' },
];

// ─── 페이지 ────────────────────────────────────────────────────
export default function DesignLabPage() {
  return (
    <>
      <Topbar title="🎨 디자인 모니터" />
      <div className="px-6 py-6 space-y-12 max-w-[1100px] mx-auto">
        {/* 헤더 안내 */}
        <div className="px-5 py-4 rounded-xl bg-neutral-900 text-white">
          <div className="flex items-center gap-2 mb-1">
            <span className="text-base font-bold">🎨 TMS 디자인 모니터</span>
            <span className="text-[10px] px-2 py-0.5 rounded bg-white/15 uppercase tracking-wider">internal tool</span>
          </div>
          <p className="text-xs opacity-80 leading-relaxed">
            사장님 + 클로드 협업 디자인 검토 도구. 향후 다른 페이지 디자인 변경 시 옵션 비교 후 결정.
            <br />
            <span className="opacity-60">운영 데이터 X · 사이드바 메뉴 미노출 · URL 직접 접근만</span>
          </p>
        </div>

        {/* ═══════════════════════════════════════════════════════════ */}
        {/* § 1. /sales 페이지 — 사장님 채택안 (2026-05-26 Phase G-4 적용 완료) */}
        {/* ═══════════════════════════════════════════════════════════ */}
        <section className="space-y-6">
          <div className="border-l-4 border-green-500 pl-4">
            <h2 className="text-lg font-bold text-indigo-black">§ 1. /sales 페이지 — 채택안 (운영 적용 완료)</h2>
            <p className="text-xs text-neutral-500 mt-1">매출 카드 = <strong className="text-green-700">안 3 (어두운 카드)</strong> · 목록 카드 = <strong className="text-green-700">안 A (좌측 색 줄 + 우측 도트)</strong></p>
          </div>

          {/* 매출 카드 */}
          <div>
            <h3 className="text-sm font-bold text-indigo-black mb-2">상단 매출 카드 (안 3 — 어두운 카드)</h3>
            <Option3Stats />
          </div>

          {/* 목록 카드 색상 규칙 */}
          <ColorLegend />

          {/* 케이스별 비교 */}
          <CaseBlock
            title="① 판매완료 — 초록색 왼편 라인 ⭐"
            description={[
              '판매 흐름이 모두 끝난 건. B2C/B2B 모두 통합 라벨.',
              '• B2C: 결제완료 + 배송완료 (delivered_at) 또는 매장수령 (송장 없음)',
              '• B2B: 출고완료 (status=\'shipped\') + 결제완료 (payment_status=\'paid\')',
              '• 배송중 → 자동 전환: ALPS cron 4시간마다',
            ]}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'paid_done').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          <CaseBlock
            title="② 진행 중 — 왼편 라인 없음 (투명)"
            description={[
              '결제완료 + 배송/출고 대기 OR 출고완료 + 결제 대기.',
              '• 배송중: 우측 ● 초록 도트',
              '• 출고 대기 / 결제 대기: 우측 ● 황색 도트',
            ]}
          >
            {SAMPLE_ROWS.filter((r) => ['paid_shipping', 'paid_wait_ship', 'shipped_b2b_unpaid'].includes(r.state)).map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          <CaseBlock
            title="③ 미결제 — 빨강색 왼편 라인 🔥"
            description={['돈 안 들어온 건. 자금 흐름 1순위 — 가장 시각적 강조.']}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'unpaid').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          <CaseBlock
            title="④ 부분결제 — 노랑색 왼편 라인"
            description={['선납금만 들어온 건. 추가 입금 받아야 할 건.']}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'partial').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          <CaseBlock
            title="⑤ 취소건 — 회색 왼편 라인 + 카드 흐림 (50%)"
            description={['취소된 건. opacity 50% + 텍스트 취소선. 사장님 시선 자동 약화.']}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'cancelled').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          {/* 전체 통합 미리보기 */}
          <div>
            <h3 className="text-sm font-bold text-indigo-black mb-2">▶ 전체 통합 미리보기 (실제 운영 시뮬레이션)</h3>
            <p className="text-xs text-neutral-500 mb-3">날짜 desc 정렬 — 사장님이 /sales 진입했을 때 보일 모습</p>
            <Wrap>
              {[...SAMPLE_ROWS].sort((a, b) => b.date.localeCompare(a.date)).map((r) => <RowA key={r.id} row={r} />)}
            </Wrap>
          </div>
        </section>

        {/* ═══════════════════════════════════════════════════════════ */}
        {/* § 2~N. 향후 다른 페이지 디자인 검토 시 여기에 § 추가 */}
        {/* ═══════════════════════════════════════════════════════════ */}
        <section className="border-2 border-dashed border-neutral-300 rounded-xl p-8 text-center bg-neutral-50">
          <div className="text-3xl mb-2">🚧</div>
          <h3 className="text-base font-bold text-neutral-700 mb-1">향후 § 추가 위치</h3>
          <p className="text-xs text-neutral-500 max-w-md mx-auto leading-relaxed">
            다른 페이지(예: /repairs, /orders, /customers, /dashboard 등) 디자인 검토 시
            <br />사장님이 클로드에게 <span className="font-semibold">"디자인 모니터에 § XXX 추가해줘"</span> 한 마디 하시면
            <br />여기에 비교 옵션 3~4개가 렌더링됩니다.
          </p>
          <div className="mt-4 text-[11px] text-neutral-400 space-y-0.5">
            <p>• 매출 카드 / 목록 카드 / 상세 패널 / 헤더 / 사이드바 등 모든 영역 가능</p>
            <p>• 비교 후 결정 → 클로드가 실제 페이지에 적용</p>
            <p>• 비교 결과는 § 1 처럼 채택안으로 박제 가능</p>
          </div>
        </section>

        {/* 푸터 */}
        <div className="text-center text-[11px] text-neutral-400 pt-4 border-t border-neutral-200">
          🎨 MAMORU TMS Design Lab · 사장님 + 클로드 협업 도구
        </div>
      </div>
    </>
  );
}

// ═══════════════════════════════════════════════════════════════
// 색상 범례 (사장님 한눈에 규칙 인지)
// ═══════════════════════════════════════════════════════════════
function ColorLegend() {
  const items = [
    { strip: 'bg-green-500', label: '판매완료', desc: 'B2C 배송완료/매장수령 · B2B 출고완료+결제완료' },
    { strip: 'bg-red-500', label: '미결제', desc: '돈 안 들어온 건 (시급)' },
    { strip: 'bg-yellow-400', label: '부분결제', desc: '선납금만 입금됨' },
    { strip: 'bg-neutral-300', label: '취소', desc: '카드 흐림 + 취소선' },
    { strip: 'bg-transparent border border-dashed border-neutral-300', label: '진행 중', desc: '결제 완료 후 배송/출고 대기, 또는 출고완료 후 결제 대기' },
  ];
  return (
    <div className="bg-white rounded-xl border border-neutral-200 p-4">
      <h3 className="text-sm font-bold text-indigo-black mb-3">🎨 색상 규칙 (좌측 4px 세로 줄)</h3>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-2">
        {items.map((it) => (
          <div key={it.label} className="flex items-center gap-3">
            <div className={`w-1 h-8 rounded ${it.strip}`} />
            <div className="flex-1">
              <div className="text-sm font-semibold text-neutral-900">{it.label}</div>
              <div className="text-xs text-neutral-500">{it.desc}</div>
            </div>
          </div>
        ))}
      </div>
      <div className="mt-4 pt-3 border-t border-neutral-100">
        <h4 className="text-xs font-bold text-neutral-700 mb-2">우측 도트 (작은 ●)</h4>
        <div className="flex flex-wrap items-center gap-4 text-xs">
          <div className="flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-green-500" /><span className="text-neutral-600">배송중 (ALPS 추적 중)</span></div>
          <div className="flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-amber-400" /><span className="text-neutral-600">출고 대기 / 결제 대기</span></div>
          <div className="flex items-center gap-2"><span className="text-neutral-400">없음 = 판매완료됐거나 미결제 등 다른 상태</span></div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// 케이스 블록 + Wrap
// ═══════════════════════════════════════════════════════════════
function CaseBlock({ title, description, children }: { title: string; description: string[]; children: React.ReactNode }) {
  return (
    <div>
      <h3 className="text-sm font-bold text-indigo-black mb-1">{title}</h3>
      <div className="text-xs text-neutral-500 mb-2 space-y-0.5">
        {description.map((d, i) => <p key={i}>{d}</p>)}
      </div>
      <Wrap>{children}</Wrap>
    </div>
  );
}
function Wrap({ children }: { children: React.ReactNode }) {
  return (
    <div className="bg-white rounded-xl border border-neutral-200 overflow-hidden">
      <div className="divide-y divide-neutral-100">{children}</div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// 매출 카드 안 3 (어두운 카드) — 채택 확정
// ═══════════════════════════════════════════════════════════════
function Option3Stats() {
  return (
    <div className="grid grid-cols-2 gap-4">
      {([
        { title: '고객 (B2C)', data: STATS.customer },
        { title: '거래처 (B2B)', data: STATS.partner },
      ] as const).map((sec) => (
        <div key={sec.title} className="bg-neutral-900 text-white rounded-xl p-5">
          <div className="text-[11px] font-semibold uppercase tracking-wider opacity-60 mb-2">{sec.title}</div>
          <div className="text-3xl font-bold tracking-tight">{formatKRW(sec.data.month.amount)}</div>
          <div className="text-xs opacity-70 mt-1">이번달 · {sec.data.month.count}건</div>
          <div className="mt-4 pt-3 border-t border-white/10 grid grid-cols-2 gap-2 text-xs">
            <div><span className="opacity-60">이번주</span><div className="font-semibold mt-0.5">{formatKRW(sec.data.week.amount)}</div></div>
            <div><span className="text-amber-300">미수금</span><div className="font-semibold mt-0.5 text-amber-300">{formatKRW(sec.data.outstanding)}</div></div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// 목록 행 — 안 A (좌측 색 줄 + 우측 도트)
// ═══════════════════════════════════════════════════════════════
function stripColor(state: RowState): string {
  switch (state) {
    case 'paid_done': return 'bg-green-500';
    case 'unpaid': return 'bg-red-500';
    case 'partial': return 'bg-yellow-400';
    case 'cancelled': return 'bg-neutral-300';
    default: return 'bg-transparent';
  }
}
function rightDot(state: RowState): { color: string; title: string } | null {
  switch (state) {
    case 'paid_shipping': return { color: 'bg-green-500', title: '배송중' };
    case 'paid_wait_ship': return { color: 'bg-amber-400', title: '출고 대기' };
    case 'shipped_b2b_unpaid': return { color: 'bg-amber-400', title: '출고완료 · 결제 대기' };
    default: return null;
  }
}
function statusLabel(state: RowState): string {
  switch (state) {
    case 'paid_done': return '판매완료';
    case 'paid_shipping': return '배송중';
    case 'paid_wait_ship': return '출고 대기';
    case 'unpaid': return '미결제';
    case 'partial': return '부분결제';
    case 'shipped_b2b_unpaid': return '출고완료 · 결제대기';
    case 'cancelled': return '취소';
  }
}
function RowA({ row }: { row: SampleRow }) {
  const isCancelled = row.state === 'cancelled';
  const strip = stripColor(row.state);
  const dot = rightDot(row.state);

  return (
    <div className={`flex items-stretch gap-3 px-4 py-3 hover:bg-warm-ivory/40 transition cursor-pointer ${isCancelled ? 'opacity-50' : ''}`}>
      <div className={`w-1 rounded ${strip}`} style={{ alignSelf: 'stretch' }} />
      <div className="flex-1 min-w-0 flex items-center gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span className={`text-sm font-semibold text-indigo-black truncate ${isCancelled ? 'line-through' : ''}`}>
              {row.name}
            </span>
            {row.note && (
              <span className="text-[10px] text-neutral-400 italic shrink-0">({row.note})</span>
            )}
          </div>
          <div className="text-xs text-neutral-500 mt-0.5 flex items-center gap-1.5 flex-wrap">
            <span>{row.date}</span>
            <span className="text-neutral-300">·</span>
            <span>{row.paymentMethod}</span>
            <span className="text-neutral-300">·</span>
            <span>{row.channelOrType}</span>
            <span className="text-neutral-300">·</span>
            <span className={
              row.state === 'unpaid' ? 'text-red-600 font-medium' :
              row.state === 'partial' ? 'text-yellow-700 font-medium' :
              row.state === 'paid_done' ? 'text-green-700 font-medium' :
              'text-neutral-500'
            }>{statusLabel(row.state)}</span>
          </div>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {dot && <span className={`w-2 h-2 rounded-full ${dot.color}`} title={dot.title} />}
          <span className={`text-sm font-bold ${isCancelled ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>
            {formatKRW(row.amount)}
          </span>
        </div>
      </div>
    </div>
  );
}
