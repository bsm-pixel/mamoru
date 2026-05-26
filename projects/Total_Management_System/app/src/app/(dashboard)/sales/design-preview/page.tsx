'use client';

/**
 * /sales 페이지 디자인 프리뷰 — 사장님 비교 선택용 (운영 X)
 *
 * 2026-05-26 Phase G-3: 매출 카드 3안 + 목록 카드 4안 비교 → 사장님 선택
 *   - 매출 카드: ✅ 안 3 (어두운 카드) 채택 (2026-05-26)
 *   - 목록 카드: 안 A 기준 상세 설계 (색상 규칙 + 케이스 확장)
 *
 * 실제 데이터 호출 X — 모두 하드코딩 임시 데이터.
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

/**
 * 목록 행 상태 분류 — 좌측 색 줄 + 우측 도트로 표현
 *
 * 색 줄 (왼쪽 4px 세로 줄):
 *   - green  = 판매완료 (B2C 배송완료/매장수령) / 정산완료 (B2B) — 모든 처리 끝남
 *   - red    = 미결제 (가장 시급)
 *   - yellow = 부분결제
 *   - gray   = 취소 + 카드 흐림
 *   - none   = 진행중 (결제완료지만 아직 배송 중 또는 정산 대기)
 *
 * 배송중 → 배송완료 자동 전환 (2026-05-25 운영 시작):
 *   ALPS 추적 cron 4시간마다 → '41'/'45' 코드 감지 → delivered_at 자동 기록
 *   → 좌측 색 줄: 투명(진행중) → 초록(판매완료) 자동 이동
 *
 * 우측 도트 (작은 ●):
 *   - green  = 배송중 / 정산 진행
 *   - amber  = 출고 대기 (송장 발급됨, 미출고)
 *   - 없음   = 도트 표시 안 함 (마무리됐거나 미결제 등 다른 상태)
 *
 * 채널/유형 표시 (작은 텍스트):
 *   - B2C: "오프라인" / "온라인상담"
 *   - B2B: "딜러" / "아카데미"
 */
type RowState = 'paid_done' | 'paid_shipping' | 'paid_wait_ship' | 'unpaid' | 'partial' | 'settled' | 'shipped_b2b' | 'cancelled';

type SampleRow = {
  id: string;
  sourceType: 'sale' | 'delivery';
  state: RowState;
  name: string;
  date: string;
  amount: number;
  paymentMethod: string;
  channelOrType: string; // "오프라인" / "온라인상담" / "딜러" / "아카데미"
  note?: string; // 케이스 설명
};

const SAMPLE_ROWS: SampleRow[] = [
  // 판매완료 / 정산완료 (초록 줄)
  { id: '1', sourceType: 'sale', state: 'paid_done', name: '서윤 점장님 (박미애)', date: '5월 26일', amount: 100000, paymentMethod: '카드', channelOrType: '온라인상담', note: '배송완료 → 판매완료 자동 전환' },
  { id: '2', sourceType: 'sale', state: 'paid_done', name: '김매장 (매장수령)', date: '5월 25일', amount: 240000, paymentMethod: '카드', channelOrType: '오프라인', note: '매장 직접수령 (배송 X, 즉시 판매완료)' },
  { id: '3', sourceType: 'delivery', state: 'settled', name: '○○매장', date: '5월 24일', amount: 1240000, paymentMethod: '이체', channelOrType: '딜러', note: 'B2B 정산완료' },

  // 진행 중 (색 줄 없음)
  { id: '4', sourceType: 'sale', state: 'paid_shipping', name: '전아연', date: '5월 25일', amount: 780000, paymentMethod: '이체', channelOrType: '온라인상담', note: '결제완료 · 배송중' },
  { id: '5', sourceType: 'sale', state: 'paid_wait_ship', name: '곽경진', date: '5월 25일', amount: 60000, paymentMethod: '이체', channelOrType: '오프라인', note: '결제완료 · 송장발급 대기' },
  { id: '6', sourceType: 'delivery', state: 'shipped_b2b', name: '△△아카데미', date: '5월 23일', amount: 580000, paymentMethod: '이체', channelOrType: '아카데미', note: '출고완료 · 정산 대기' },

  // 미수금 (빨강 줄)
  { id: '7', sourceType: 'sale', state: 'unpaid', name: '김승한', date: '5월 18일', amount: 280000, paymentMethod: '카드', channelOrType: '오프라인', note: '미결제 — 가장 시급' },
  { id: '8', sourceType: 'delivery', state: 'unpaid', name: '□□아카데미', date: '5월 20일', amount: 1500000, paymentMethod: '이체', channelOrType: '아카데미', note: '거래처 미결제 (큰 금액)' },

  // 부분결제 (노랑 줄)
  { id: '9', sourceType: 'sale', state: 'partial', name: '박상민', date: '5월 22일', amount: 350000, paymentMethod: '카드', channelOrType: '온라인상담', note: '부분결제 (선납금만)' },

  // 취소 (회색 줄)
  { id: '10', sourceType: 'sale', state: 'cancelled', name: '취소건 예시', date: '5월 20일', amount: 150000, paymentMethod: '카드', channelOrType: '오프라인', note: '취소 — 흐림 표시' },
];

// ─── 페이지 ────────────────────────────────────────────────────
export default function SalesDesignPreviewPage() {
  return (
    <>
      <Topbar title="디자인 프리뷰 (사장님 선택용)" />
      <div className="px-6 py-6 space-y-10 max-w-[1100px] mx-auto">
        {/* 안내 */}
        <div className="px-4 py-3 rounded-lg bg-yellow-50 border border-yellow-200 text-sm text-yellow-800">
          💡 <strong>디자인 프리뷰</strong> — 운영 데이터 X, UI 미리보기만.
          <span className="ml-2 text-xs">매출 카드 = ✅ <strong>안 3 채택</strong> / 목록 카드 = <strong>안 A 상세 설계</strong> 진행 중</span>
        </div>

        {/* § A 매출 카드 — 안 3 채택 확정 */}
        <section>
          <h2 className="text-base font-bold text-indigo-black mb-3">
            § A. 상단 매출 카드 — ✅ <span className="text-green-600">안 3 채택</span>
          </h2>
          <Option3Stats />
          <p className="mt-2 text-xs text-neutral-500">사장님 결정 (2026-05-26): 어두운 카드 + 화이트 텍스트. /sales 메인 페이지 Phase G-4 에서 이 디자인 적용 예정.</p>
        </section>

        {/* § B 목록 카드 — 안 A 상세 설계 */}
        <section className="space-y-6">
          <h2 className="text-base font-bold text-indigo-black">
            § B. 좌측 목록 카드 — <span className="text-green-600">안 A 기준 상세 설계</span>
          </h2>

          {/* 색상 규칙 범례 */}
          <ColorLegend />

          {/* 케이스 1 — 판매완료 / 정산완료 (초록 줄) */}
          <CaseBlock
            title="① 판매완료 · 정산완료 — 초록색 왼편 라인 ⭐"
            description={[
              '판매·납품 흐름이 모두 끝난 건. 사장님이 "이건 다 됐다" 인식할 1순위.',
              '• B2C 판매: 결제완료 + 배송완료 (delivered_at 채워짐) → "판매완료"',
              '• B2C 매장 직접수령: 결제완료 + 송장 없음 (즉시 완료) → "판매완료"',
              '• B2B 거래처: 결제완료 + 정산완료 (status=\'settled\') → "정산완료"',
              '• 배송중 → 배송완료 자동 전환: ALPS cron 4시간마다 자동 추적 (운영 검증 완료)',
            ]}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'paid_done' || r.state === 'settled').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          {/* 케이스 2 — 진행중 (색 줄 없음) */}
          <CaseBlock
            title="② 진행 중 — 왼편 라인 없음 (투명)"
            description={[
              '결제는 완료됐지만 후속 처리가 남은 건. 사장님이 "이건 처리해야 한다" 인식.',
              '• 배송중: 송장 발급 + ALPS 추적 중 (우측 ● 초록 도트)',
              '• 출고 대기: 송장 발급됐지만 미출고 (우측 ● 황색 도트)',
              '• B2B 출고완료 · 정산 대기: 거래처 납품 후 settle 전',
            ]}
          >
            {SAMPLE_ROWS.filter((r) => ['paid_shipping', 'paid_wait_ship', 'shipped_b2b'].includes(r.state)).map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          {/* 케이스 3 — 미결제 (빨강 줄) */}
          <CaseBlock
            title="③ 미결제 — 빨강색 왼편 라인 🔥"
            description={[
              '돈 안 들어온 건. 사장님 자금 흐름 1순위 — 가장 시각적으로 강조.',
              '• B2C 미결제 (외상 거래)',
              '• B2B 거래처 미결제 (큰 금액 — 사장님이 한눈에 인지해야 함)',
            ]}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'unpaid').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          {/* 케이스 4 — 부분결제 (노랑 줄) */}
          <CaseBlock
            title="④ 부분결제 — 노랑색 왼편 라인"
            description={[
              '선납금만 들어온 건. 사장님 후속 추가 입금 받아야 할 건.',
              '• 결제수단 옆에 "부분결제" 표기 + 노랑 라인',
            ]}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'partial').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          {/* 케이스 5 — 취소 (회색 줄 + 흐림) */}
          <CaseBlock
            title="⑤ 취소건 — 회색 왼편 라인 + 카드 흐림 (50%)"
            description={[
              '취소된 건. 시각적으로 약화 (opacity 50%) + 텍스트 취소선',
              '• 카드 자체가 흐려져서 사장님 시선이 안 감 — 정상',
            ]}
          >
            {SAMPLE_ROWS.filter((r) => r.state === 'cancelled').map((r) => <RowA key={r.id} row={r} />)}
          </CaseBlock>

          {/* 전체 통합 미리보기 — 실제 운영 화면 시뮬레이션 */}
          <section>
            <h3 className="text-sm font-bold text-indigo-black mb-2">▶ 전체 통합 미리보기 (실제 운영 화면 시뮬레이션)</h3>
            <p className="text-xs text-neutral-500 mb-3">날짜 desc 정렬 — 사장님이 /sales 페이지에 진입했을 때 보일 모습</p>
            <Wrap>
              {[...SAMPLE_ROWS].sort((a, b) => b.date.localeCompare(a.date)).map((r) => <RowA key={r.id} row={r} />)}
            </Wrap>
          </section>
        </section>

        {/* 추가 메모 */}
        <section className="px-4 py-3 rounded-lg bg-neutral-50 border border-neutral-200 text-xs text-neutral-600 space-y-1">
          <p className="font-bold text-neutral-900 mb-1">📝 사장님 추가 변경 요청 시 알려주세요:</p>
          <p>• 색상 규칙 조정 (어떤 상태를 어떤 색으로?)</p>
          <p>• 우측 도트 표시 케이스 추가/제거</p>
          <p>• 채널/유형 표기 방식 (텍스트 vs 아이콘 vs 작은 칩)</p>
          <p>• 폰트 크기 / 간격 / 호버 효과</p>
        </section>
      </div>
    </>
  );
}

// ═══════════════════════════════════════════════════════════════
// 색상 범례 (사장님 한눈에 규칙 인지)
// ═══════════════════════════════════════════════════════════════
function ColorLegend() {
  const items = [
    { strip: 'bg-green-500', label: '판매완료 · 정산완료', desc: 'B2C 배송완료/매장수령 · B2B 정산완료' },
    { strip: 'bg-red-500', label: '미결제', desc: '돈 안 들어온 건 (시급)' },
    { strip: 'bg-yellow-400', label: '부분결제', desc: '선납금만 입금됨' },
    { strip: 'bg-neutral-300', label: '취소', desc: '카드 흐림 + 취소선' },
    { strip: 'bg-transparent border border-dashed border-neutral-300', label: '진행 중', desc: '결제 완료 + 후속 처리 중 (라인 없음)' },
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
          <div className="flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-green-500" /><span className="text-neutral-600">배송중 / 정산 진행</span></div>
          <div className="flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-amber-400" /><span className="text-neutral-600">출고 대기</span></div>
          <div className="flex items-center gap-2"><span className="text-neutral-400">없음 = 마무리됐거나 미결제 등 다른 상태</span></div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// 케이스 블록 (제목 + 설명 + 행 리스트)
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
// 목록 행 — 안 A 상세 (색 줄 + 도트 + 채널 + 케이스 노트)
// ═══════════════════════════════════════════════════════════════

function stripColor(state: RowState): string {
  switch (state) {
    case 'paid_done':
    case 'settled':
      return 'bg-green-500'; // 마무리 완료
    case 'unpaid':
      return 'bg-red-500'; // 미결제
    case 'partial':
      return 'bg-yellow-400'; // 부분결제
    case 'cancelled':
      return 'bg-neutral-300'; // 취소
    default:
      return 'bg-transparent'; // 진행 중
  }
}

function rightDot(state: RowState): { color: string; title: string } | null {
  switch (state) {
    case 'paid_shipping':
      return { color: 'bg-green-500', title: '배송중' };
    case 'paid_wait_ship':
      return { color: 'bg-amber-400', title: '출고 대기' };
    case 'shipped_b2b':
      return { color: 'bg-green-500', title: '출고완료 · 정산 대기' };
    default:
      return null;
  }
}

function statusLabel(state: RowState): string {
  switch (state) {
    case 'paid_done': return '판매완료';
    case 'paid_shipping': return '배송중';
    case 'paid_wait_ship': return '출고 대기';
    case 'unpaid': return '미결제';
    case 'partial': return '부분결제';
    case 'settled': return '정산완료';
    case 'shipped_b2b': return '출고완료';
    case 'cancelled': return '취소';
  }
}

function RowA({ row }: { row: SampleRow }) {
  const isCancelled = row.state === 'cancelled';
  const strip = stripColor(row.state);
  const dot = rightDot(row.state);

  return (
    <div className={`flex items-stretch gap-3 px-4 py-3 hover:bg-warm-ivory/40 transition cursor-pointer ${isCancelled ? 'opacity-50' : ''}`}>
      {/* 좌측 색 줄 */}
      <div className={`w-1 rounded ${strip}`} style={{ alignSelf: 'stretch' }} />

      {/* 본문 */}
      <div className="flex-1 min-w-0 flex items-center gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span className={`text-sm font-semibold text-indigo-black truncate ${isCancelled ? 'line-through' : ''}`}>
              {row.name}
            </span>
            {/* 케이스 노트 — 디자인 프리뷰 전용 (실제 운영 페이지에는 표시 안 함) */}
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
              row.state === 'paid_done' || row.state === 'settled' ? 'text-green-700 font-medium' :
              'text-neutral-500'
            }>{statusLabel(row.state)}</span>
          </div>
        </div>

        {/* 우측: 도트 + 금액 */}
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
