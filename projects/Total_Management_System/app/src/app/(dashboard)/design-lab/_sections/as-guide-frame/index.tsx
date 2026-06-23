'use client';

/**
 * § 복원수리 안내 페이지 — 개선 틀 (2026-06-20)
 * IA 재배치 와이어프레임 + 변경 요약 + 회사소개 역할분담.
 * 운영 적용 후 이 섹션 삭제(디자인 모니터 회전 룰).
 */

type Kind = 'keep' | 'move' | 'new' | 'fill';

const KIND_STYLE: Record<Kind, { box: string; tag: string; label: string }> = {
  keep: { box: 'border-stone-300 bg-white', tag: 'bg-stone-200 text-stone-600', label: '기존' },
  move: { box: 'border-blue-400 bg-blue-50', tag: 'bg-blue-600 text-white', label: '↑ 이동' },
  new:  { box: 'border-emerald-400 bg-emerald-50', tag: 'bg-emerald-600 text-white', label: '🆕 신규' },
  fill: { box: 'border-amber-400 bg-amber-50 border-dashed', tag: 'bg-amber-500 text-white', label: '채움 필요' },
};

type Block = { n: string; title: string; sub?: string; kind: Kind };

const PROPOSED: Block[] = [
  { n: '1', title: '히어로 영상', sub: '작업 장면 클로즈업 (루프)', kind: 'keep' },
  { n: '2', title: '선언', sub: '"새 가위를 처음 만났던 그 순간으로, 복원합니다"', kind: 'keep' },
  { n: '3', title: '⭐ Before & After', sub: '결과 증거(영상+사진) — 상단으로 끌어올림', kind: 'move' },
  { n: '4', title: 'PREVIEW', sub: '작업 과정 클립 (과정 티저)', kind: 'keep' },
  { n: '5', title: '마모루의 원칙 3', sub: '책임 · 진단 · 투명', kind: 'keep' },
  { n: '6', title: '수냉식 신념 + 대표', sub: '백성민 — 대표 사진 채우기', kind: 'fill' },
  { n: '7', title: '복원 후기 (미용사 보이스)', sub: '"첫 커트감이 돌아왔다" 1~2줄 — 전환 ↑', kind: 'new' },
  { n: '8', title: '수리내역서 샘플', sub: '투명함을 "보여주기" — 실물 1장', kind: 'new' },
  { n: '9', title: '접수 CTA', sub: '복원수리 접수하기', kind: 'keep' },
];

const CURRENT_ORDER = [
  '히어로영상', '선언', 'CTA', '원칙3', '작업사진 갤러리',
  'PREVIEW', 'Before&After', '신념', '대표', '9단계 공정(이미지 비어있음)',
];

function FrameBlock({ b }: { b: Block }) {
  const s = KIND_STYLE[b.kind];
  return (
    <div className={`rounded-lg border px-3 py-2.5 ${s.box}`}>
      <div className="flex items-center gap-1.5">
        <span className="w-5 h-5 shrink-0 rounded bg-stone-900 text-white text-[10px] font-bold flex items-center justify-center">{b.n}</span>
        <span className="text-[13px] font-bold text-stone-900 flex-1">{b.title}</span>
        <span className={`text-[9px] px-1.5 py-0.5 rounded font-bold shrink-0 ${s.tag}`}>{s.label}</span>
      </div>
      {b.sub && <p className="text-[11px] text-stone-500 mt-1 pl-[26px] leading-snug">{b.sub}</p>}
    </div>
  );
}

export function AsGuideFrameSection() {
  return (
    <section className="rounded-2xl border border-stone-200 bg-white p-5">
      <div className="flex items-center gap-2 mb-1">
        <h2 className="text-base font-bold text-stone-900">§ 복원수리 안내 페이지 — 개선 틀</h2>
        <span className="text-[10px] px-2 py-0.5 rounded bg-stone-100 text-stone-500">2026-06-20 · 와이어프레임</span>
      </div>
      <p className="text-xs text-stone-500 mb-5">
        intro 탭(다크 랜딩) 제안 구조입니다. 색 = 변경 종류. 운영 적용 후 이 섹션은 비웁니다.
      </p>

      <div className="grid lg:grid-cols-[300px_1fr] gap-6">
        {/* 모바일 프레임 — 제안 구조 */}
        <div>
          <p className="text-[11px] font-bold text-stone-500 mb-2">제안 구조 (모바일)</p>
          <div className="mx-auto w-[290px] rounded-[26px] border-[6px] border-stone-900 bg-stone-900 p-2 shadow-xl">
            <div className="rounded-[18px] bg-stone-100 overflow-hidden">
              <div className="bg-stone-900 text-white text-center text-[10px] py-1.5 tracking-widest">MAMORU · 복원수리 안내</div>
              <div className="p-2.5 space-y-2 max-h-[560px] overflow-y-auto">
                {PROPOSED.map((b) => <FrameBlock key={b.n} b={b} />)}
                <div className="rounded-lg border border-stone-300 bg-stone-50 px-3 py-2">
                  <p className="text-[10px] font-bold text-stone-400 uppercase tracking-wider mb-1">하단 탭</p>
                  <div className="flex flex-wrap gap-1">
                    {['과정안내', '소요시간', '비용안내', '포장방법', 'QnA'].map((t) => (
                      <span key={t} className="text-[10px] px-2 py-0.5 rounded-full bg-white border border-stone-200 text-stone-600">{t}</span>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </div>
          {/* 범례 */}
          <div className="flex flex-wrap gap-2 justify-center mt-3">
            {(['keep', 'move', 'new', 'fill'] as Kind[]).map((k) => (
              <span key={k} className={`text-[10px] px-2 py-0.5 rounded font-bold ${KIND_STYLE[k].tag}`}>{KIND_STYLE[k].label}</span>
            ))}
          </div>
        </div>

        {/* 변경 요약 + 회사소개 역할분담 */}
        <div className="space-y-5">
          <div>
            <p className="text-[11px] font-bold text-stone-500 mb-2">핵심 변경 5</p>
            <ul className="space-y-1.5 text-[13px] text-stone-700">
              <li>① <b>증거(Before&After)를 상단으로</b> — 결과부터 보여 신뢰 선점 (현재는 한참 아래)</li>
              <li>② <b>복원 후기 신규</b> — 미용사 보이스 1~2줄(현재 0건, 전환 손해)</li>
              <li>③ <b>수리내역서 샘플 신규</b> — "투명함"을 글이 아닌 실물로 증명</li>
              <li>④ <b>대표 사진 채움</b> — 책임자 얼굴 = 신뢰 핵심인데 비어있음</li>
              <li>⑤ <b>9단계 공정</b> 이미지 9개 비어있음 → 채우거나 임시 숨김(미완성 노출 제거)</li>
            </ul>
          </div>

          <div>
            <p className="text-[11px] font-bold text-stone-500 mb-2">현재 순서 (참고)</p>
            <div className="flex flex-wrap gap-1">
              {CURRENT_ORDER.map((t, i) => (
                <span key={i} className={`text-[10px] px-2 py-0.5 rounded ${t.includes('비어') ? 'bg-amber-100 text-amber-700' : 'bg-stone-100 text-stone-500'}`}>{i + 1}.{t}</span>
              ))}
            </div>
          </div>

          <div>
            <p className="text-[11px] font-bold text-stone-500 mb-2">회사소개(메인) ↔ 안내 — 중복은 OK, 역할만 분리</p>
            <div className="rounded-lg border border-stone-200 overflow-hidden text-[12px]">
              <div className="grid grid-cols-[90px_1fr] bg-stone-50 border-b border-stone-200">
                <div className="px-3 py-2 font-bold text-stone-600">메인</div>
                <div className="px-3 py-2 text-stone-700">한 줄 <b>약속</b> — "직접·2대·첫 커트감" 압축 + 링크</div>
              </div>
              <div className="grid grid-cols-[90px_1fr]">
                <div className="px-3 py-2 font-bold text-stone-600 border-r border-stone-100">안내</div>
                <div className="px-3 py-2 text-stone-700">그 약속의 <b>증명</b> — 원칙·공정·증거·후기·비용·CTA</div>
              </div>
            </div>
            <p className="text-[11px] text-stone-500 mt-2">
              핵심 기둥(2대 기술 / 직접 복원 / 첫 커트감 / 투명)은 <b>반복해야</b> 각인됩니다. 메인=압축, 안내=증명.
              <br />⚠️ proof 키워드 통일 필요: 메인 "일본 공장 정밀" vs 안내 "수냉식" → 대표 무기 하나로 정리.
            </p>
          </div>
        </div>
      </div>
    </section>
  );
}
